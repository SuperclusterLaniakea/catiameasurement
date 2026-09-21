#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
控制 CATIA 测量零件/装配体并导出 Excel:
  体积、表面积、质量(重量)、包络体(包围盒)长宽高、用户参数、三维标注

用法:
  python measure_catia.py                      # 测量, 生成 测量结果.txt / 测量结果.xlsx / results.json
  python measure_catia.py --density 2.70       # 指定密度 (g/cm3, 默认 1.0)
  python measure_catia.py --dir test           # 指定目录 (默认 test)
  python measure_catia.py --shots              # 三维标注逐个截图并写入 Excel (需先运行一次主程序生成 results.json)
  python measure_catia.py --no-part-shot       # 仅测几何/参数, 不截整机图(加快批量)

说明:
  - 通过 COM 自动化连接 CATIA(未运行则自动启动), 测量为只读操作
  - 体积/表面积: SPAWorkbench.Measurable (m3/m2); 质量 = 体积 x 密度
  - 包络体: 优先用 SPA Measurable.GetBoundingBox(按几何体并集, 更可靠),
    不可用时回退"远置平面法":在 ±1e6 mm 处创建 6 个临时平面, 用
    GetMinimumDistance 反推各轴向极值得到 AABB 包围盒, 测完删除临时几何体
  - 异常包围盒: 某方向尺寸过大且远大于其余两维时标记为疑似异常(多为并集进了
    远置几何体), 在报告/Excel 备注中提示"仅作参考"
  - 用户参数: 参数名中只含一个反斜杠(顶层, PartName\\ParamName)的视为用户参数
  - 几何图形集参数: 挂在"几何图形集(图形数据集合)"节点上的参数, 判据为参数全名去掉
    零件名与末段参数名后, 剩余路径恰好等于某几何图形集的层级路径(见 hybrid_body_paths)
  - 整机截图(默认开启): 测量每个零件时, 单独窗口显示该零件 -> FitAllIn 居中
    -> 隐藏结构树(SpecificationTreeActivation) -> 白背景(SetWindowBackgroundColor)
    -> CATIA 内置 Viewer.CaptureToFile 截 800x600 图, 写入 <目录>/零件截图/,
    并嵌入 Excel "测量结果" 页该零件所在行的"截图"列(同一行)
  - 三维标注: 通过 part.AnnotationSets 枚举; 选中标注后用同样视图截图写回 Excel
  - Excel: 页1 测量结果(每零件一行含整机截图), 页2 用户参数, 页3 参数汇总, 页4 三维标注
"""
import argparse
import json
import os
import re
import sys
import time
import tempfile

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

import pythoncom
import win32com.client
import win32com.client.dynamic  # late-bound 派发: 运行期经 IDispatch 解析 .Part/.Product
import win32con

# 强制 win32com 对所有 COM 对象(含方法返回的 Document/Part/Product 等)使用 late-bound 包裹。
# 否则 gen_py 早期绑定下基类 Document 不含 .Part/.Product, 任何 doc.Part/doc.Product 都会抛
# "'Document' object has no attribute 'Part'"。让 gencache 返回 None 即退化成 dynamic 包裹。
import win32com.client.gencache as _gencache
_gencache.GetClassForCLSID = lambda *a, **k: None
_gencache.GetClassForProgID = lambda *a, **k: None
import win32gui
import win32process
import threading
from PIL import Image, ImageFilter, ImageChops
import xlsxwriter

DEFAULT_DENSITY = 1.0  # g/cm3, 默认密度
PLANE_LIMIT = 1000000.0  # 远置平面距离 (mm), 零件须位于 ±1km 内
PART_SHOT_SIZE = (800, 600)  # 截图时 CATIA 窗口/图片尺寸(像素), 白背景无结构树
SHOT_ROW_HEIGHT = 60    # 截图在 Excel 中的显示高度 = 所在行行高(点), 与行高一致
MBD_SHOT_ROW_HEIGHT = 450  # MBD 捕获截图行高(点): 800x600 像素按 96DPI -> 600*0.75=450pt
FORCE_CAPTURE = False   # 由 --force 置位: 忽略已存在的截图缓存, 强制重新生成


def log(msg):
    print(msg, flush=True)


def close_catia_dialogs():
    """关闭 CATIA 进程内残留的模态对话框(打开/超级输入消息等), 防止阻塞 COM"""
    pid = None

    def cb(h, _):
        nonlocal pid
        try:
            t = win32gui.GetWindowText(h)
            if pid is None and "CATIA V5" in t:
                _, pid = win32process.GetWindowThreadProcessId(h)
            if pid is not None and win32gui.IsWindowVisible(h):
                t2 = win32gui.GetWindowText(h)
                _, p = win32process.GetWindowThreadProcessId(h)
                if p == pid and ("打开" in t2 or "超级输入" in t2):
                    win32gui.PostMessage(h, win32con.WM_CLOSE, 0, 0)
        except Exception:
            pass
        return True

    try:
        win32gui.EnumWindows(cb, None)
    except Exception:
        pass


def connect_catia():
    """连接 CATIA, 未运行则启动并等待就绪。
    必须用 late-bound(dynamic) dispatch: 否则 gen_py 早期绑定下 Document 基类不含
    .Part/.Product, 会报 "'Document' object has no attribute 'Part'"。dynamic 在运行期
    经 IDispatch 解析 .Part/.Product, 对 PartDocument/ProductDocument 均生效。"""
    close_catia_dialogs()
    try:
        ole = pythoncom.GetActiveObject("CATIA.Application")
        catia = win32com.client.dynamic.Dispatch(ole)
        log("已连接正在运行的 CATIA")
    except Exception:
        log("CATIA 未运行, 正在启动(首次启动可能需要 1~2 分钟)...")
        catia = win32com.client.dynamic.Dispatch("CATIA.Application")
        ready = False
        for _ in range(120):  # 最多等 4 分钟
            try:
                _ = catia.Documents.Count
                ready = True
                break
            except Exception:
                time.sleep(2)
        if not ready:
            raise RuntimeError("CATIA 启动超时, 请确认已安装并完成许可配置")
        log("CATIA 已就绪")
    try:
        catia.DisplayFileAlerts = False  # 关闭文件提示对话框
    except Exception:
        pass
    return catia


def _open_document(catia, path, timeout=180):
    """打开文档并自动处理 CATIA 的"定位引用文档"对话框。

    很多零件含外部参考(如紧固件库 hlv_*.CATPart), 本机缺失这些文件时,
    CATIA 会弹出"定位引用文档"(标题"打开")模态对话框并阻塞 Documents.Open。
    该对话框的"关闭"按钮会以"参考未解析"方式继续打开文档——零件自身实体几何
    仍然完整(可测量、可截图)。本函数在调用线程执行 Open, 另起守护线程监视并
    自动点击"关闭", 避免批量测量时人工介入卡死。"""
    stop = threading.Event()

    def _monitor():
        elapsed = 0.0
        while not stop.is_set() and elapsed < timeout:
            time.sleep(0.4)
            elapsed += 0.4
            try:
                wins = []

                def _enum(h, a):
                    if win32gui.IsWindowVisible(h) and win32gui.GetWindowText(h).strip():
                        a.append((h, win32gui.GetWindowText(h)))

                win32gui.EnumWindows(_enum, wins)
                for h, t in wins:
                    ts = t.strip()
                    # CATIA "定位引用文档"对话框: 标题为 打开/查找/定位/Locate
                    if ts in ("打开", "查找", "定位", "Locate", "Open") or ts.endswith("打开"):
                        kids = []

                        def _kid(kh, ka):
                            if win32gui.IsWindowVisible(kh):
                                ka.append((kh, win32gui.GetWindowText(kh)))

                        win32gui.EnumChildWindows(h, _kid, kids)
                        clicked = False
                        for kh, kt in kids:
                            if kt.strip() == "关闭":
                                win32gui.SendMessage(kh, win32con.BM_CLICK, 0, 0)
                                clicked = True
                                break
                        if not clicked:
                            # 兜底: 直接关闭对话框窗口
                            win32gui.PostMessage(h, win32con.WM_CLOSE, 0, 0)
                        # 继续监视, 以防出现多个定位对话框
            except Exception:
                pass

    th = threading.Thread(target=_monitor, daemon=True)
    th.start()
    try:
        doc = catia.Documents.Open(path)
    finally:
        stop.set()
        th.join(timeout=3)
    return doc



def auto_density(doc):
    """尝试读取零件材质密度(kg/m3 -> g/cm3)。
    无材质时 CATIA 的 Inertia.Density 返回默认 1000 kg/m3, 以及哨兵值
    1 / -1, 均不可信, 仅当 >1000 时才认为是真实材质密度。"""
    try:
        product = doc.Product
        inertia = product.GetTechnologicalObject("Inertia")
        d = float(inertia.Density)  # kg/m3
        if d > 1000.0:
            return d / 1000.0
    except Exception:
        pass
    return None


def collect_user_params(part):
    """收集用户参数: 顶层参数(名称只含一个反斜杠, PartName\\ParamName)。
    系统参数通常嵌套较深(如 ...\\Axis Systems\\...\\X), 顶层即用户定义参数。
    返回 [(名称, 值字符串), ...]"""
    out = []
    try:
        params = part.Parameters
        for i in range(1, params.Count + 1):
            try:
                p = params.Item(i)
                nm = str(p.Name)
                if nm.count("\\") == 1:
                    val = ""
                    try:
                        val = str(p.ValueAsString())  # 方法, 通常含单位
                    except Exception:
                        try:
                            val = str(p.Value)
                        except Exception:
                            pass
                    out.append((nm.split("\\")[-1], val))
            except Exception:
                continue
    except Exception:
        pass
    return out


def hybrid_body_paths(part):
    """递归收集零件所有"几何图形集"(Geometrical Set / 图形数据集合)的层级路径。
    返回 set, 元素如 'Part Notes:'、'Goemetry'、'Goemetry\\Planes'。"""
    paths = set()

    def walk(hbs, prefix):
        try:
            for i in range(1, hbs.Count + 1):
                b = hbs.Item(i)
                p = prefix + [str(b.Name)]
                paths.add("\\".join(p))
                try:
                    walk(b.HybridBodies, p)
                except Exception:
                    pass
        except Exception:
            pass

    try:
        walk(part.HybridBodies, [])
    except Exception:
        pass
    return paths


def collect_geomset_params(part):
    """提取挂在"几何图形集"(Geometrical Set / 图形数据集合)**节点本身**上的参数与参数值。

    BUG 根因(已修): 旧实现按参数名里含 'Geometrical Sets'/'几何图形集' 匹配, 而实际几何图形集
    几乎都是自定义名(Part Notes: / Standard Notes: / Material Description: / ECCN / Goemetry /
    BONDING...), 故历史上一律匹配 0 条、图形数据集合里的参数全部丢失。

    正确判据: CATIA 参数全名形如  PartName\\<路径...>\\<参数名>  —— 去掉末段(参数名)与首段
    (零件名)后, 若剩余路径**恰好是某个几何图形集的层级路径**, 则该参数是直接挂在该图形集
    节点上的(如 '...\\Part Notes:\\5SN0000007687')。
    而 '...\\Goemetry\\Planes\\平面.49\\偏移' 这种中间还夹着特征名(平面.49)的属于特征内部
    参数(活动/偏移/模式等系统噪声), 不计入。

    返回 [(显示名, 值字符串), ...], 显示名='几何图形集名\\参数名'。"""
    out = []
    try:
        gp_set = hybrid_body_paths(part)
        if not gp_set:
            return out
        params = part.Parameters
        for i in range(1, params.Count + 1):
            try:
                p = params.Item(i)
                nm = str(p.Name)
                parts = nm.split("\\")
                if len(parts) < 3:
                    continue
                owner = "\\".join(parts[1:-1])      # 去掉零件名与末段参数名
                if owner not in gp_set:
                    continue
                try:
                    val = str(p.ValueAsString())    # 方法, 通常含单位
                except Exception:
                    try:
                        val = str(p.Value)
                    except Exception:
                        val = ""
                out.append((owner + "\\" + parts[-1], val))
            except Exception:
                continue
    except Exception:
        pass
    return out


def collect_annotations(part):
    """收集三维标注: 遍历标注集, 返回 [(标注集, 标注名, 标注文本, 标注对象), ...]"""
    out = []
    try:
        sets = part.AnnotationSets
    except Exception:
        return out
    for si in range(1, sets.Count + 1):
        try:
            s = sets.Item(si)
        except Exception:
            continue
        set_name = ""
        try:
            set_name = str(s.Name)
        except Exception:
            pass
        try:
            anns = s.Annotations
        except Exception:
            continue
        for ai in range(1, anns.Count + 1):
            try:
                a = anns.Item(ai)
            except Exception:
                continue
            ann_name = ""
            try:
                ann_name = str(a.Name)
            except Exception:
                pass
            ann_text = ""
            try:
                ann_text = str(a.Text)
            except Exception:
                pass
            out.append((set_name, ann_name, ann_text, a))
    return out


def sanitize_fname(name):
    """清理文件名非法字符"""
    return re.sub(r'[\\/:*?"<>|\r\n\t]', "_", name).strip()


def measure_part(doc, density_g_cm3):
    """测量零件: 体积(m3)/表面积(m2)/质量(kg)/包络体(mm)"""
    part = doc.Part
    spa = doc.GetWorkbench("SPAWorkbench")

    vol_m3 = 0.0
    area_m2 = 0.0
    try:
        ref = part.CreateReferenceFromObject(part)
        m = spa.GetMeasurable(ref)
        vol_m3 = float(m.Volume) or 0.0
        area_m2 = float(m.Area) or 0.0
    except Exception:
        vol_m3 = 0.0
        area_m2 = 0.0

    if vol_m3 <= 0.0:
        vol_m3 = 0.0
        area_m2 = 0.0
        bodies = part.Bodies
        for bi in range(1, bodies.Count + 1):
            try:
                body = bodies.Item(bi)
                ref = part.CreateReferenceFromObject(body)
                m = spa.GetMeasurable(ref)
                v = float(m.Volume) or 0.0
                a = float(m.Area) or 0.0
                vol_m3 += v
                area_m2 += a
            except Exception:
                continue

    density_used = density_g_cm3
    density_src = "未检测到零件材质, 按参数密度"
    ad = auto_density(doc)
    if ad is not None:
        density_used = ad
        density_src = "零件材质(自动检测)"

    mass_kg = vol_m3 * density_used * 1000.0
    weight_n = mass_kg * 9.80665

    box = part_bbox(part, spa, doc)  # (L, W, H) mm 或 None
    try:
        bodies_count = part.Bodies.Count
    except Exception:
        bodies_count = 1

    # 异常包围盒标记: 某方向尺寸过大且远大于其余两维, 多半是并集进了远置几何体
    box_suspect = False
    if box is not None:
        L, W, H = box
        mx = max(L, W, H)
        if mx > 5000.0 and mx > 10.0 * (min(L, W, H) + 1e-9):
            box_suspect = True

    return {
        "vol_m3": vol_m3,
        "area_m2": area_m2,
        "density_g_cm3": density_used,
        "density_src": density_src,
        "mass_kg": mass_kg,
        "weight_n": weight_n,
        "box_mm": box,
        "box_suspect": box_suspect,
        "bodies_count": bodies_count,
    }


def _length_factor(doc):
    """文档长度单位 -> 毫米 的换算系数(用于包围盒尺寸)。
    读取失败或未知单位时按毫米处理。枚举值: mm=1, m=2, cm=3, ft=4, in=5。"""
    try:
        u = doc.Units
        L = int(u.Length)
        return {1: 1.0, 2: 1000.0, 3: 10.0, 4: 304.8, 5: 25.4}.get(L, 1.0)
    except Exception:
        return 1.0


def part_bbox_getbbox(part, doc):
    """计算与坐标轴对齐的 AABB 包围盒, 返回 (L,W,H) mm 或 None。
    方法: 在 ±PLANE_LIMIT 处创建 6 个临时平行平面, 用 SPA Measurable.
    GetMinimumDistance 逐几何体(Body)反推各轴向极值得到该几何体包围盒, 再对
    保留的几何体取并集。测量后删除临时几何体, 不保存, 不改动零件。
    注意: 本机 CATIA V5(B32) 的 Measurable.GetBoundingBox 对单几何体(Body)
    调用失败(返回尺寸<=0), 故包络体只用远置平面法。
    远置辅助体(如 *_502_RH_ 右手件里被放在 ~24000mm 处的镜像参考体)会把并集
    尺寸异常撑大; 通过"某几何体最大边长 >> 其余几何体最大边长(>5倍)"识别并排除,
    得到反映真实零件的包络体。若排除后无保留体或任一方向尺寸过大则返回 None。"""
    try:
        spa = doc.GetWorkbench("SPAWorkbench")
    except Exception:
        return None
    temp = None
    try:
        hsf = part.HybridShapeFactory
        temp = part.HybridBodies.Add()
        temp.Name = "TmpBBoxMeasure"
        oe = part.OriginElements
        bases = [("x", oe.PlaneYZ), ("y", oe.PlaneZX), ("z", oe.PlaneXY)]
        planes = {}
        for axis, base in bases:
            bref = part.CreateReferenceFromObject(base)
            # iOrientation: 0 -> +LIM 侧(沿法向), 1 -> -LIM 侧(实测确认)
            planes[(axis, "pos")] = hsf.AddNewPlaneOffset(bref, PLANE_LIMIT, 0)
            planes[(axis, "neg")] = hsf.AddNewPlaneOffset(bref, PLANE_LIMIT, 1)
            temp.AppendHybridShape(planes[(axis, "pos")])
            temp.AppendHybridShape(planes[(axis, "neg")])
        # 只更新临时包围盒体(依赖原点元素, 不牵动断裂的外部参考),
        # 避免对含缺失外部参考的零件执行整件 Update 时卡死
        try:
            part.UpdateObject(temp)
        except Exception:
            pass
        bodies = part.Bodies
        n = bodies.Count
        if n == 0:
            return None
        boxes = []  # 每个几何体: (minx,maxx,miny,maxy,minz,maxz) mm
        for bi in range(1, n + 1):
            try:
                body = bodies.Item(bi)
                m = spa.GetMeasurable(part.CreateReferenceFromObject(body))
                lo = {}
                hi = {}
                for axis in "xyz":
                    d_pos = float(m.GetMinimumDistance(part.CreateReferenceFromObject(planes[(axis, "pos")])))
                    d_neg = float(m.GetMinimumDistance(part.CreateReferenceFromObject(planes[(axis, "neg")])))
                    lo[axis] = d_neg - PLANE_LIMIT
                    hi[axis] = PLANE_LIMIT - d_pos
                boxes.append((lo["x"], hi["x"], lo["y"], hi["y"], lo["z"], hi["z"]))
            except Exception:
                continue
        if not boxes:
            return None
        # 排除远置辅助体(如 *_RH_ 件的镜像副本被放在对称远处): 迭代地移除"离当前
        # 保留体包围盒中心质心最远、且距离远超其自身最大边长(>5倍)"的几何体, 直到
        # 无可移除或只剩一个。这样 RH 件会收敛到单一真实实体, 不影响正常多实体零件
        # (本项目中所有零件整体被平移到了距原点约 33000mm 处, 故不能用"离原点"判据)。
        centers = [((b[0] + b[1]) / 2.0, (b[2] + b[3]) / 2.0, (b[4] + b[5]) / 2.0) for b in boxes]
        dims = [max(b[1] - b[0], b[3] - b[2], b[5] - b[4]) for b in boxes]
        keep_idx = list(range(len(boxes)))
        while len(keep_idx) > 1:
            cs = [centers[k] for k in keep_idx]
            cxm = sum(c[0] for c in cs) / len(cs)
            cym = sum(c[1] for c in cs) / len(cs)
            czm = sum(c[2] for c in cs) / len(cs)
            far = -1; fard = -1.0
            for j, k in enumerate(keep_idx):
                d = ((centers[k][0] - cxm) ** 2 + (centers[k][1] - cym) ** 2 + (centers[k][2] - czm) ** 2) ** 0.5
                if d > fard:
                    fard = d; far = j
            if dims[keep_idx[far]] > 0 and fard > 5.0 * dims[keep_idx[far]]:
                keep_idx.pop(far)
            else:
                break
        keep = [boxes[k] for k in keep_idx]
        if len(keep) < len(boxes):
            print("[包络体] 排除 {} 个远置/镜像几何体, 保留 {}".format(len(boxes) - len(keep), len(keep)), flush=True)
        minx = min(b[0] for b in keep); maxx = max(b[1] for b in keep)
        miny = min(b[2] for b in keep); maxy = max(b[3] for b in keep)
        minz = min(b[4] for b in keep); maxz = max(b[5] for b in keep)
        L = maxx - minx; W = maxy - miny; H = maxz - minz
        if L <= 0 or W <= 0 or H <= 0 or max(L, W, H) > 2 * PLANE_LIMIT:
            return None
        return (L, W, H)
    except Exception:
        return None
    finally:
        if temp is not None:
            try:
                sel = doc.Selection
                sel.Clear()
                sel.Add(temp)
                sel.Delete()
            except Exception:
                pass


def part_bbox(part, spa, doc):
    """计算与坐标轴对齐的 AABB 包围盒, 返回 (L,W,H) mm 或 None。
    调用 part_bbox_getbbox 完成实际测量(远置平面法逐几何体 + 排除远置辅助体);
    仅当该路径也失败时才无条件回退到简单的远置平面法并集。
    测量后删除临时几何体, 不保存, 不改动零件。"""
    # 优先用 GetBoundingBox(更可靠); 失败再回退远置平面法
    gb = part_bbox_getbbox(part, doc)
    if gb is not None:
        return gb
    temp = None
    try:
        hsf = part.HybridShapeFactory
        temp = part.HybridBodies.Add()
        temp.Name = "TmpBBoxMeasure"
        oe = part.OriginElements
        bases = [("x", oe.PlaneYZ), ("y", oe.PlaneZX), ("z", oe.PlaneXY)]
        planes = {}
        for axis, base in bases:
            bref = part.CreateReferenceFromObject(base)
            # iOrientation: 0 -> +LIM 侧(沿法向), 1 -> -LIM 侧(实测确认)
            planes[(axis, "pos")] = hsf.AddNewPlaneOffset(bref, PLANE_LIMIT, 0)
            planes[(axis, "neg")] = hsf.AddNewPlaneOffset(bref, PLANE_LIMIT, 1)
            temp.AppendHybridShape(planes[(axis, "pos")])
            temp.AppendHybridShape(planes[(axis, "neg")])
        # 只更新临时包围盒体(依赖原点元素, 不牵动断裂的外部参考),
        # 避免对含缺失外部参考的零件执行整件 Update 时卡死
        try:
            part.UpdateObject(temp)
        except Exception:
            pass

        min_p = {"x": -PLANE_LIMIT, "y": -PLANE_LIMIT, "z": -PLANE_LIMIT}
        max_p = {"x": PLANE_LIMIT, "y": PLANE_LIMIT, "z": PLANE_LIMIT}
        got = False
        bodies = part.Bodies
        for bi in range(1, bodies.Count + 1):
            try:
                body = bodies.Item(bi)
                m = spa.GetMeasurable(part.CreateReferenceFromObject(body))
                lo = {}
                hi = {}
                for axis in "xyz":
                    d_pos = float(m.GetMinimumDistance(part.CreateReferenceFromObject(planes[(axis, "pos")])))
                    d_neg = float(m.GetMinimumDistance(part.CreateReferenceFromObject(planes[(axis, "neg")])))
                    lo[axis] = d_neg - PLANE_LIMIT
                    hi[axis] = PLANE_LIMIT - d_pos
                if not got:
                    min_p = dict(lo)
                    max_p = dict(hi)
                    got = True
                else:
                    for axis in "xyz":
                        min_p[axis] = min(min_p[axis], lo[axis])
                        max_p[axis] = max(max_p[axis], hi[axis])
            except Exception:
                continue
        if not got:
            return None
        L = max_p["x"] - min_p["x"]
        W = max_p["y"] - min_p["y"]
        H = max_p["z"] - min_p["z"]
        if L <= 0 or W <= 0 or H <= 0 or max(L, W, H) > 2 * PLANE_LIMIT:
            return None
        return (L, W, H)
    except Exception:
        return None
    finally:
        if temp is not None:
            try:
                sel = doc.Selection
                sel.Clear()
                sel.Add(temp)
                sel.Delete()
            except Exception:
                pass


def count_products(product):
    """统计装配体子项数量(递归)"""
    total = 0
    try:
        prods = product.Products
    except Exception:
        return 0
    for i in range(1, prods.Count + 1):
        try:
            p = prods.Item(i)
        except Exception:
            continue
        total += 1
        try:
            if p.Products.Count > 0:
                total += count_products(p)
        except Exception:
            pass
    return total


def measure_product(doc, density_g_cm3):
    """测量装配体: 递归统计零件实例, 体积/质量按零件求和"""
    product = doc.Product
    total_vol = 0.0
    instances = 0
    unique = set()
    leaves = 0

    def walk(prod):
        nonlocal total_vol, instances, leaves
        try:
            children = prod.Products
            n = children.Count
        except Exception:
            n = 0
        if n == 0:
            leaves += 1
            try:
                pd = prod.ReferenceProduct.Parent  # PartDocument
                p = pd.Part
                spa = pd.GetWorkbench("SPAWorkbench")
                ref = p.CreateReferenceFromObject(p)
                m = spa.GetMeasurable(ref)
                v = float(m.Volume) or 0.0
                total_vol += v
                instances += 1
                try:
                    unique.add(pd.Name)
                except Exception:
                    pass
            except Exception:
                pass
        else:
            for i in range(1, n + 1):
                try:
                    walk(children.Item(i))
                except Exception:
                    pass

    walk(product)

    density_used = density_g_cm3
    density_src = "未检测到材质, 按参数密度"
    ad = auto_density(doc)
    if ad is not None:
        density_used = ad
        density_src = "零件材质(自动检测)"

    mass_kg = total_vol * density_used * 1000.0
    return {
        "vol_m3": total_vol,
        "mass_kg": mass_kg,
        "instances": instances,
        "unique_parts": len(unique),
        "leaves": leaves,
        "density_g_cm3": density_used,
        "density_src": density_src,
    }


def fmt_box(box):
    if box is None:
        return "无法计算"
    l, w, h = box
    return "长 {:.3f} mm x 宽 {:.3f} mm x 高 {:.3f} mm".format(l, w, h)


def detect_type(doc, fname):
    """判定文档类型: 零件(Part) 或 装配体(Product)。
    注意: CATIA 自动化的 Document 没有 Type 属性, 不能使用 doc.Type。"""
    ext = os.path.splitext(fname)[1].lower()
    if ext == ".catpart":
        return "Part"
    if ext == ".catproduct":
        return "Product"
    try:
        doc.Part
        return "Part"
    except Exception:
        return "Product"


def measure_one(catia, path, density_g_cm3, opts=None):
    """测量单个文件, 返回 (ok, 文本报告, 结构化数据)。
    opts 字典控制提取内容(缺省全开):
      measure 几何测量(体积/表面积/质量/重量/包络体)
      user_params 用户参数, annotations 三维标注, part_shot 零件整机截图"""
    opts = opts or {}
    do_measure = opts.get("measure", True)
    do_params = opts.get("user_params", True)
    do_anns = opts.get("annotations", True)
    do_shot = opts.get("part_shot", True)  # 默认生成整机截图, 写入 Excel 测量结果行的"截图"列(同一行)
    # MBD 导出(移植自 MBD.txt VBA): 指定参数 / Capture 截图(写入 Word 报告)
    do_mbd_params = opts.get("mbd_params", False)
    do_mbd_caps = opts.get("mbd_captures", False)
    fname = os.path.basename(path)
    doc = None
    data = {
        "key": fname, "file": fname, "kind": "Part", "ok": False, "error": "",
        "vol_m3": 0.0, "area_m2": 0.0, "density_g_cm3": density_g_cm3,
        "density_src": "", "mass_kg": 0.0, "weight_n": 0.0,
        "box": None, "box_suspect": False, "note": "", "user_params": [], "annotations": [],
        "part_shot": "",
        "mbd_captures": 0, "mbd_params": [], "mbd_capture_shots": {},
    }
    try:
        log("  打开: " + fname)
        # 设置文件搜索目录: 让 CATIA 在打开装配体(.CATProduct)时,
        # 能自动按目录加载同目录下的子 CATPart(否则易报"零件文件缺失/未加载")
        try:
            catia.FileSearchDirectories = os.path.dirname(path)
        except Exception:
            pass
        doc = _open_document(catia, path)
        typ = detect_type(doc, fname)
        data["kind"] = typ
        lines = []
        lines.append("=" * 60)
        lines.append("文件: " + fname)
        lines.append("类型: " + ("零件" if typ == "Part" else "装配体"))
        if not do_measure:
            lines.append("几何测量: 未勾选(跳过)")

        if typ == "Part":
            r = None
            if do_measure:
                r = measure_part(doc, density_g_cm3)
                data.update(r)
                data["box"] = list(r["box_mm"]) if r["box_mm"] else None
            up = []
            gp = []
            if do_params:
                up = collect_user_params(part=doc.Part)
                data["user_params"] = up
                # "几何图形集"(图形数据集合)节点上的参数, 单独成表(值多为长文本, 不并入用户参数)
                gp = collect_geomset_params(doc.Part)
                data["geomset_params"] = gp
            if do_anns:
                anns = collect_annotations(doc.Part)
                data["annotations"] = [(s, n, t) for s, n, t, _ in anns]
            # MBD 提取(源自 MBD.txt VBA 移植): 指定参数 + Capture 截图, 写入同一个 Excel
            # 注意: 必须在整机截图(do_shot)之前执行 —— capture_part 会把 viewer 背景改成白色,
            # 而 Capture 截图要求 CATIA 原始的蓝色渐变背景, 顺序颠倒会得到白底截图。
            note_parts = []
            if do_mbd_params or do_mbd_caps:
                try:
                    shots_dir_m = os.path.join(os.path.dirname(path), "MBD截图")
                    os.makedirs(shots_dir_m, exist_ok=True)
                    log("  MBD 提取: 参数={} Capture截图={}".format(do_mbd_params, do_mbd_caps))
                    if do_mbd_params:
                        mbd_rows = collect_mbd_params(doc.Part)
                        data["mbd_params"] = mbd_rows
                        lines.append("MBD 指定参数({}): {}".format(
                            len(mbd_rows), " / ".join(
                                "{}={}".format(dsp, v) for _, dsp, v in mbd_rows[:20])))
                    if do_mbd_caps:
                        caps_map = {}
                        n_caps = 0
                        base_key = fname[:fname.rfind(".")] if "." in fname else fname
                        try:
                            sets = doc.Part.AnnotationSets
                            for si in range(1, sets.Count + 1):
                                try:
                                    caps = sets.Item(si).Captures
                                except Exception:
                                    continue
                                for j in range(1, caps.Count + 1):
                                    try:
                                        cap = caps.Item(j)
                                    except Exception:
                                        continue
                                    cap_name = ""
                                    try:
                                        cap_name = str(cap.Name)
                                    except Exception:
                                        pass
                                    if not cap_name:
                                        cap_name = "Capture_{}".format(n_caps + 1)
                                    log("    Capture 截图: " + cap_name)
                                    img = capture_capture_view(catia, cap, n_caps + 1, shots_dir_m)
                                    n_caps += 1
                                    if img:
                                        # 截图永久保留, 以中文捕获名为文件名, 写入 Excel
                                        img_permanent = os.path.join(
                                            shots_dir_m, "{}_{}.jpg".format(
                                                sanitize_fname(base_key), sanitize_fname(cap_name)))
                                        try:
                                            os.replace(img, img_permanent)
                                            img = img_permanent
                                        except Exception:
                                            pass
                                        caps_map[cap_name] = img
                                    try:
                                        doc.Selection.Clear()
                                    except Exception:
                                        pass
                        except Exception:
                            pass
                        data["mbd_captures"] = n_caps
                        data["mbd_capture_shots"] = caps_map
                        lines.append("MBD Capture 截图: {} 张 -> {}".format(len(caps_map), shots_dir_m))
                    note_parts.append("MBD 已提取(参数 {}, Capture 截图 {})".format(
                        len(data.get("mbd_params") or []),
                        len(data.get("mbd_capture_shots") or {})))
                except Exception as e:
                    lines.append("MBD 提取异常: " + repr(e)[:120])

            if do_shot:
                # 零件整机截图, 供 Excel 测量结果行的"截图"列使用
                # (放在 MBD 之后: MBD Capture 截图需保留原始蓝色背景,
                #  而 capture_part 会把背景强制改为白色, 顺序不能颠倒)
                try:
                    part_shots_dir = os.path.join(os.path.dirname(path), "零件截图")
                    data["part_shot"] = capture_part(catia, doc, part_shots_dir, fname) or ""
                except Exception:
                    data["part_shot"] = ""

            if r is not None:
                lines.append("体积  : {:.6f} m3  ( {:.3f} cm3 )".format(r["vol_m3"], r["vol_m3"] * 1e6))
                lines.append("表面积: {:.6f} m2  ( {:.3f} cm2 )".format(r["area_m2"], r["area_m2"] * 1e4))
                lines.append("密度  : {:.3f} g/cm3 ({})".format(r["density_g_cm3"], r["density_src"]))
                lines.append("质量  : {:.4f} kg  ( {:.3f} g )".format(r["mass_kg"], r["mass_kg"] * 1000.0))
                lines.append("重量  : {:.3f} N".format(r["weight_n"]))
                lines.append("包络体: " + fmt_box(r["box_mm"]) + (
                    "  (含 {} 个几何体, 为并集)".format(r["bodies_count"]) if r["bodies_count"] > 1 else ""
                ))
                if r["bodies_count"] > 1:
                    note_parts.append("含 {} 个几何体(包络体为并集)".format(r["bodies_count"]))
                if r.get("box_suspect"):
                    note_parts.append("包络体疑似异常(含远置几何体), 仅作参考")
            if up:
                lines.append("用户参数({}): {}".format(
                    len(up), " / ".join("{}={}".format(n, v) for n, v in up)))
            if gp:
                lines.append("几何图形集参数({}): {}".format(
                    len(gp), " / ".join("{}={}".format(n, v[:60]) for n, v in gp)))
            if data["annotations"]:
                lines.append("三维标注({}): {}".format(
                    len(data["annotations"]), " / ".join(n for _, n, _ in data["annotations"])))
            if data["part_shot"]:
                note_parts.append("整机截图 1 张")
            note_parts.append("三维标注 {} 个; 用户参数 {} 个".format(len(data["annotations"]), len(up)))
            data["note"] = "; ".join(note_parts)
            ok = (r is not None and (r["vol_m3"] > 0 or r["box_mm"] is not None)) if do_measure else True
        else:
            if do_measure:
                r = measure_product(doc, density_g_cm3)
                data.update(r)
                data["weight_n"] = r["mass_kg"] * 9.80665
                lines.append("装配体总体积: {:.6f} m3  ( {:.3f} cm3 )".format(r["vol_m3"], r["vol_m3"] * 1e6))
                if r["instances"] > 0:
                    lines.append("装配体总质量: {:.4f} kg  (体积和 x 密度 {:.3f} g/cm3, {})".format(
                        r["mass_kg"], r["density_g_cm3"], r["density_src"]))
                    lines.append("零件实例数: {}  (独立零件: {})".format(r["instances"], r["unique_parts"]))
                    data["note"] = "零件实例 {} 个(独立零件 {})".format(r["instances"], r["unique_parts"])
                else:
                    lines.append("装配体总质量: 无法获取")
                    if r["leaves"] > 0:
                        lines.append("  装配体含 {} 个子项, 但零件几何无法解析(零件文件缺失或未加载)".format(r["leaves"]))
                        data["note"] = "含 {} 个子项, 零件文件缺失或未加载".format(r["leaves"])
                    else:
                        lines.append("  装配体为空(不含子项)")
                        data["note"] = "装配体为空"
                    try:
                        inertia = doc.Product.GetTechnologicalObject("Inertia")
                        lines.append("  (Inertia 参考质量: {:.4f} kg)".format(float(inertia.Mass)))
                    except Exception:
                        pass
                lines.append("包络体: 装配体级未计算(可对内部零件逐一测量)")
            ok = True

        lines.append("=" * 60)
        data["ok"] = ok
        return ok, "\n".join(lines) + "\n", data
    except Exception as e:
        data["ok"] = False
        data["error"] = str(e)
        return False, "文件 {} 处理失败: {}\n".format(fname, e), data
    finally:
        if doc is not None:
            try:
                doc.Close()
            except Exception:
                pass


# ============================ 三维标注截图 ============================

def capture_window(hwnd, path):
    """用 PrintWindow 捕获窗口内容(先强制重绘, PrintWindow flag 2->1->0 降级尝试)。
    CATIA 的 3D 视口是 OpenGL, 不同显卡/驱动对 PrintWindow flag 兼容性不同,
    逐个尝试以提高抓到 OpenGL 渲染内容的成功率。返回是否成功。"""
    import ctypes
    import win32ui
    user32 = ctypes.windll.user32
    hwndDC = None
    try:
        user32.PrintWindow.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_uint32]
        user32.PrintWindow.restype = ctypes.c_bool
        try:
            # RDW_INVALIDATE | RDW_ALLCHILDREN | RDW_UPDATENOW: 先强制窗口重绘, 再捕获
            user32.RedrawWindow(hwnd, None, None, 0x0001 | 0x0080 | 0x0100)
        except Exception:
            pass
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        w, h = right - left, bottom - top
        if w <= 0 or h <= 0:
            return False
        hwndDC = win32gui.GetWindowDC(hwnd)
        mfcDC = win32ui.CreateDCFromHandle(hwndDC)
        saveDC = mfcDC.CreateCompatibleDC()
        bitmap = win32ui.CreateBitmap()
        bitmap.CreateCompatibleBitmap(mfcDC, w, h)
        saveDC.SelectObject(bitmap)
        ok = False
        for flag in (2, 1, 0):  # PW_RENDERFULLCONTENT -> 普通 -> 兼容模式
            ok = bool(user32.PrintWindow(hwnd, saveDC.GetSafeHdc(), flag))
            if ok:
                break
        bmpstr = bitmap.GetBitmapBits(True)
        img = Image.frombuffer("RGB", (w, h), bmpstr, "raw", "BGRX", 0, 1)
        img.save(path)
        mfcDC.DeleteDC()
        saveDC.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwndDC)
        hwndDC = None
        return ok
    except Exception:
        if hwndDC is not None:
            try:
                win32gui.ReleaseDC(hwnd, hwndDC)
            except Exception:
                pass
        return False


def find_catia_window(base):
    """查找 CATIA 中显示该零件的窗口句柄。
    优先取前台窗口(若属于 CATIA), 失败则枚举所有窗口, 仅匹配 CATIA 拥有的窗口
    (标题含 'CATIA V5' / '.CATPart' / '.CATProduct'), 先按文件名精确匹配,
    再退化为文件名(去扩展名)包含匹配。返回 HWND 或 None。"""
    hwnd = None
    try:
        fg = win32gui.GetForegroundWindow()
        if fg:
            t = win32gui.GetWindowText(fg)
            if "CATIA V5" in t or ".CATPart" in t or ".CATProduct" in t:
                if base in t:
                    hwnd = fg
    except Exception:
        pass
    if hwnd is not None:
        return hwnd

    exact = None
    fuzzy = None

    def find(h, _):
        nonlocal exact, fuzzy
        try:
            t = win32gui.GetWindowText(h)
            if "CATIA V5" in t or ".CATPart" in t or ".CATProduct" in t:
                if base in t:
                    # 精确(文件名含匹配); 若有 '.CATPart'/'CATProduct' 后缀则优先
                    if exact is None:
                        exact = h
                elif fuzzy is None:
                    fuzzy = h
        except Exception:
            pass
        return True

    try:
        win32gui.EnumWindows(find, None)
    except Exception:
        pass
    return exact or fuzzy


def _is_target_size(path):
    """判断已有截图是否为目标尺寸(800x600), 用于断点续跑时识别旧尺寸截图"""
    try:
        with Image.open(path) as im:
            return im.size == PART_SHOT_SIZE
    except Exception:
        return False


def _image_has_content(path):
    """粗略检测截图是否纯色/空白: 灰度方差过小(近似纯色)视为无内容, 截图失败"""
    try:
        with Image.open(path) as im:
            im2 = im.convert("L").resize((100, 75))
            px = list(im2.getdata())
            mean = sum(px) / len(px)
            var = sum((v - mean) ** 2 for v in px) / len(px)
            return var > 3.0
    except Exception:
        return True


def _image_score(path):
    """给截图打分(多帧择优): 中央视口区域"白色背景占比 x 灰度方差"。
    纯白(白背景已生效但模型未渲染, 方差≈0)与灰色 UI(无白背景, 白色占比≈0)
    都得低分; 白底+模型(白色占比高且内容丰富)得高分。"""
    try:
        with Image.open(path) as im:
            g = im.convert("L")
            w, h = g.size
            box = g.crop((int(w * 0.25), int(h * 0.2), int(w * 0.75), int(h * 0.8)))
            px = list(box.getdata())
            n = len(px)
            white = sum(1 for v in px if v > 230) / n   # 白背景占比
            mean = sum(px) / n
            var = sum((v - mean) ** 2 for v in px) / n  # 内容量(模型边缘等)
            return white * var
    except Exception:
        return -1.0


def _viewer_capture(viewer, tmp):
    """用 CATIA 软件内的图像捕获工具(Viewer.CaptureToFile)截图:
    直接渲染当前视口到文件, 不受屏幕抓取/OpenGL 兼容性影响。
    CATIA 的 CatCaptureFormat 枚举: BMP=0, JPEG=1, TIFF=2, PNG=3, EMF=4。
    格式降级尝试: PNG(3) -> BMP(0) -> JPEG(1), 统一转成 PNG 存入 tmp。
    返回是否成功。"""
    if viewer is None:
        return False
    base = os.path.splitext(tmp)[0]
    # 正确的枚举顺序: PNG=3 优先, 其次 BMP=0, 最后 JPEG=1
    for fmt, ext in ((3, ".png"), (0, ".bmp"), (1, ".jpg")):
        out = base + ext
        try:
            viewer.CaptureToFile(fmt, out)
            if os.path.exists(out) and os.path.getsize(out) > 0:
                with Image.open(out) as im:
                    im.save(tmp)
                try:
                    os.remove(out)
                except Exception:
                    pass
                return True
        except Exception:
            try:
                if os.path.exists(out):
                    os.remove(out)
            except Exception:
                pass
    return False


def _to_target_size(tmp, path):
    """把临时截图统一为 800x600(保持纵横比, 白底填充, 配合白色背景), 保存到 path。
    注意: 必须在关闭 Image.open 的文件句柄后再改名/删除, 否则 Windows
    下文件被占用(共享冲突)导致失败; 失败时直接把临时图改名过去。"""
    try:
        with Image.open(tmp) as img:
            w, h = img.size
            same_size = (w, h) == PART_SHOT_SIZE
            if not same_size:
                scale = min(PART_SHOT_SIZE[0] / w, PART_SHOT_SIZE[1] / h)
                nw, nh = max(1, int(round(w * scale))), max(1, int(round(h * scale)))
                img2 = img.resize((nw, nh), Image.LANCZOS)
                canvas = Image.new("RGB", PART_SHOT_SIZE, (255, 255, 255))  # 白色背景
                canvas.paste(img2, ((PART_SHOT_SIZE[0] - nw) // 2, (PART_SHOT_SIZE[1] - nh) // 2))
                canvas.save(path)
        # 退出 with 后句柄已释放, 此时改名/删除不会冲突
        if same_size:
            os.replace(tmp, path)
        else:
            try:
                os.remove(tmp)
            except Exception:
                pass
    except Exception:
        try:
            os.replace(tmp, path)
        except Exception:
            pass


def _hide_tree(catia, hwnd):
    """用 F3 快捷键隐藏结构树(best-effort)。把本线程输入附加到 CATIA 窗口线程,
    置前台后发送 F3, 再解除附加。任何失败静默忽略(不影响截图主流程)。"""
    cat_tid = None
    try:
        import win32api, win32process
        our_tid = win32api.GetCurrentThreadId()
        cat_tid, _ = win32process.GetWindowThreadProcessId(hwnd)
        if cat_tid and cat_tid != our_tid:
            win32process.AttachThreadInput(our_tid, cat_tid, True)
        try:
            win32gui.SetForegroundWindow(hwnd)
        except Exception:
            pass
        time.sleep(0.3)
        win32api.keybd_event(0x72, 0, 0, 0)   # F3 down
        win32api.keybd_event(0x72, 0, 2, 0)   # F3 up
    except Exception:
        pass
    finally:
        try:
            if cat_tid:
                import win32api, win32process
                our_tid = win32api.GetCurrentThreadId()
                win32process.AttachThreadInput(our_tid, cat_tid, False)
        except Exception:
            pass


def _show_tree(catia, hwnd):
    """恢复结构树显示(再发一次 F3)。"""
    _hide_tree(catia, hwnd)


def _whiten_bg(raw, dst, size=PART_SHOT_SIZE, tol=40):
    """把 CaptureToFile 抓到的 3D 视图处理为纯白底, 并缩放到 size(默认 800x600) 居中补白。
    背景判定(只抠背景, 绝不误删零件):
      - 近白色像素(三通道均 >=237) -> 直接置白;
      - 渐变背景(当左右端点构成浅色渐变时)且与像素偏差<tol -> 置白, 但要求该渐变本身为浅色
        (min 通道 >=200), 防止把浅色零件(如浅蓝 (210,210,255)) 误判为背景而删除。
    结构树/罗盘/坐标轴由 _isolate_part 单独处理, 这里只管背景。返回是否成功。"""
    try:
        with Image.open(raw) as im:
            im = im.convert("RGB")
        w, h = im.size
        px = im.load()
        out = im.copy()
        opx = out.load()
        CL = px[3, h // 2]
        CR = px[w - 4, h // 2]
        for x in range(w):
            t = x / (w - 1)
            gr = int(CL[0] + (CR[0] - CL[0]) * t)
            gg = int(CL[1] + (CR[1] - CL[1]) * t)
            gb = int(CL[2] + (CR[2] - CL[2]) * t)
            grad_light = (min(gr, gg, gb) >= 200)  # 渐变本身是浅色(白/浅蓝)才当背景
            for y in range(h):
                r, g, b = px[x, y]
                if r >= 237 and g >= 237 and b >= 237:
                    opx[x, y] = (255, 255, 255)
                elif (grad_light and abs(r - gr) < tol
                      and abs(g - gg) < tol and abs(b - gb) < tol):
                    opx[x, y] = (255, 255, 255)
        # 缩放补白到目标尺寸
        tw, th = size
        scale = min(tw / w, th / h)
        nw, nh = max(1, int(w * scale)), max(1, int(h * scale))
        out = out.resize((nw, nh), Image.LANCZOS)
        canvas = Image.new("RGB", (tw, th), (255, 255, 255))
        canvas.paste(out, ((tw - nw) // 2, (th - nh) // 2))
        canvas.save(dst)
        return True
    except Exception:
        return False


def _isolate_part(im, min_wh=16, margin=6, white_thr=240):
    """从整幅 3D 截图中"抠出零件本体", 其余(结构树文字/图标/竖条、罗盘、坐标轴、背景)
    全部置白。返回 (处理后图像, 零件边界框(x0,y0,x1,y1))。

    原理: 对"非白像素"做连通域分析(8 连通)。零件是"边界框面积最大"的连通块; 若零件由
    多个几何体组成, 则再保留与零件边界框相交的"较大"连通块(宽高均>=min_wh)。结构树
    文字/图标/竖线、罗盘、坐标轴都是细小连通块, 被一并清除。
    实测本机 CATIA 无法用 COM/F3 隐藏结构树, 且树与 3D 视图同属一个画布, 故只能在
    截图上按此方法剔除。"""
    w, h = im.size
    if w < 40 or h < 40:
        return im, (0, 0, w - 1, h - 1)
    g = im.convert("L")
    bw = g.point(lambda v: 255 if v < white_thr else 0)
    px = bw.load()
    parent = []
    all_runs = []

    def _find(a):
        while parent[a] != a:
            parent[a] = parent[parent[a]]
            a = parent[a]
        return a

    def _union(a, b):
        ra, rb = _find(a), _find(b)
        if ra != rb:
            parent[rb] = ra

    runs_prev = []
    for y in range(h):
        x = 0
        runs = []
        while x < w:
            if px[x, y]:
                x0 = x
                while x < w and px[x, y]:
                    x += 1
                runs.append((x0, x - 1))
            else:
                x += 1
        cur = []
        for (x0, x1) in runs:
            node = len(parent)
            parent.append(node)
            all_runs.append((y, x0, x1, node))
            for (qx0, qx1, qn) in runs_prev:
                if not (x1 < qx0 - 1 or x0 > qx1 + 1):   # 8 连通(允许相邻)
                    _union(node, qn)
            cur.append((x0, x1, node))
        runs_prev = cur
    if not all_runs:
        return im, (0, 0, w - 1, h - 1)

    bbox = {}
    pixcnt = {}
    for (y, x0, x1, node) in all_runs:
        r = _find(node)
        b = bbox.get(r)
        if b is None:
            bbox[r] = [x0, x1, y, y]
            pixcnt[r] = x1 - x0 + 1
        else:
            if x0 < b[0]:
                b[0] = x0
            if x1 > b[1]:
                b[1] = x1
            if y < b[2]:
                b[2] = y
            if y > b[3]:
                b[3] = y
            pixcnt[r] += x1 - x0 + 1

    # 零件 = 边界框面积最大的连通块; 先排除"竖直细条"(结构树左侧竖条: 窄且近贯穿全高)
    def _area(r):
        b = bbox[r]
        return (b[1] - b[0] + 1) * (b[3] - b[2] + 1)

    def _is_tall_bar(r):
        b = bbox[r]
        return (b[1] - b[0] + 1) < 0.10 * w and (b[3] - b[2] + 1) > 0.70 * h

    cands = [r for r in bbox if not _is_tall_bar(r)]
    if not cands:
        cands = list(bbox)
    part_root = max(cands, key=_area)
    pb = bbox[part_root]
    part_pix = pixcnt[part_root]
    # 额外保留(零件由多个几何体组成时)的门槛: 尺寸足够大且像素数达零件 5% 以上,
    # 借此排除结构树文字/图标(细小块)与罗盘/坐标轴。
    min_pix = max(800, int(0.05 * part_pix))
    keep = {part_root}
    for r, b in bbox.items():
        if r == part_root:
            continue
        bw_ = b[1] - b[0] + 1
        bh_ = b[3] - b[2] + 1
        if bw_ < min_wh or bh_ < min_wh or pixcnt.get(r, 0) < min_pix:
            continue
        if (b[1] < pb[0] - margin or b[0] > pb[1] + margin
                or b[3] < pb[2] - margin or b[2] > pb[3] + margin):
            continue   # 与零件相距较远(树/罗盘/坐标轴) -> 丢弃
        keep.add(r)
    mask = Image.new("L", (w, h), 0)
    mpx = mask.load()
    for (y, x0, x1, node) in all_runs:
        if _find(node) in keep:
            for x in range(x0, x1 + 1):
                mpx[x, y] = 255
    white = Image.new("RGB", im.size, (255, 255, 255))
    out = Image.composite(im, white, mask)
    return out, (pb[0], pb[2], pb[1], pb[3])


def _fit_center_on_white(im, bbox, size, pad_ratio=0.08):
    """把图像裁剪到零件 bbox(留 pad_ratio 边距)并居中缩放到 size, 白底填充。"""
    w, h = im.size
    x0, y0, x1, y1 = bbox
    x0 = max(0, min(x0, w - 1))
    x1 = max(0, min(x1, w - 1))
    y0 = max(0, min(y0, h - 1))
    y1 = max(0, min(y1, h - 1))
    if x1 <= x0 or y1 <= y0:
        return im
    mw = int((x1 - x0 + 1) * pad_ratio) + 4
    mh = int((y1 - y0 + 1) * pad_ratio) + 4
    crop = im.crop((max(0, x0 - mw), max(0, y0 - mh),
                    min(w, x1 + 1 + mw), min(h, y1 + 1 + mh)))
    tw, th = size
    cw, ch = crop.size
    if cw <= 0 or ch <= 0:
        return im
    scale = min(tw / cw, th / ch)
    nw, nh = max(1, int(cw * scale)), max(1, int(ch * scale))
    crop = crop.resize((nw, nh), Image.LANCZOS)
    canvas = Image.new("RGB", (tw, th), (255, 255, 255))
    canvas.paste(crop, ((tw - nw) // 2, (th - nh) // 2))
    return canvas


def _capture_view(catia, doc, hwnd, tmp, base, fit_center=False, select=None, frames=3):
    """按顺序准备视图并截图:
       1) 窗口置顶可见 + 激活文档, 设白色背景, Reframe 居中
       2) 预热(丢弃首帧, 避免黑帧污染后续统计)
       3) RenderingMode=6(Shading with Edges) 实体着色 —— 实测本机 CATIA 的 2/3/4 是
          线框/轮廓(无表面色), 只有 5/6/7 才是实体着色, 6 最接近 CATIA 默认显示
       4) 强制重绘(RefreshDisplay 关->开) + 多帧捕获取 _image_score 最优
       5) 后处理: 连通域抠出零件(去结构树/罗盘/坐标轴/背景) -> 裁剪居中 -> 800x600
    返回 True 表示 tmp 已写入有效截图。"""
    # 1) 窗口置顶可见
    try:
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
    except Exception:
        pass
    try:
        win32gui.SetWindowPos(hwnd, win32con.HWND_TOPMOST, 40, 40,
                              PART_SHOT_SIZE[0], PART_SHOT_SIZE[1],
                              win32con.SWP_SHOWWINDOW)
    except Exception:
        pass
    try:
        doc.Activate()
    except Exception:
        pass
    # 取 viewer(正确路径: ActiveWindow.ActiveViewer / ActiveWindow.Viewers.Item(1))
    viewer = None
    for getter in (lambda: catia.ActiveWindow.ActiveViewer,
                  lambda: catia.ActiveWindow.Viewers.Item(1)):
        try:
            v = getter()
            if v is not None:
                viewer = v
                break
        except Exception:
            pass

    def _refresh():
        try:
            catia.RefreshDisplay = False
        except Exception:
            pass
        try:
            catia.RefreshDisplay = True
        except Exception:
            pass

    def _set_mode(m):
        if viewer is not None:
            try:
                viewer.RenderingMode = m
            except Exception:
                pass

    # --- 先设白背景 + 居中 + 预热: 保证后续探树/截图拿到"已重绘的白底帧"。
    #     关键! 探树若在设白背景前进行, 会抓到深色(近黑)背景帧 -> 深色像素统计失真。 ---
    if viewer is not None:
        try:
            viewer.PutBackgroundColor([1.0, 1.0, 1.0])
        except Exception:
            pass
    if fit_center and viewer is not None:
        try:
            viewer.Reframe()
        except Exception:
            pass
    # 选中标注(如有)
    sel = None
    if select:
        try:
            sel = doc.Selection
            sel.Clear()
            for a in select:
                try:
                    sel.Add(a)
                except Exception:
                    pass
        except Exception:
            sel = None
    # 预热: 线框模式 + 强制重绘 + 丢弃首帧(首帧常为黑/未重绘, 会污染树检测)
    _set_mode(2)
    _refresh()
    time.sleep(0.7)
    raw_path = os.path.join(tempfile.gettempdir(), "catia_cap_raw.png")   # 单一路径复用: 每帧覆写, 不放工程目录、不做批量删除
    if viewer is not None:
        try:
            viewer.CaptureToFile(2, raw_path)
        except Exception:
            pass

    # 1.5) 隐藏结构树: 树与 3D 视图共用同一画布, 树可见时其节点标签的白色背景框会盖住
    #      零件形成"白洞"——该白洞与背景同色, 连通域后处理无法恢复, 必须真正隐藏结构树。
    #      实测: 后台进程发 F3 完全无效(窗口无法取得前台, 键事件送不进 CATIA), 已弃用;
    #      viewer.FullScreen=True 是唯一可靠的 COM 途径, 会隐藏结构树+工具条, 全窗口显示 3D。
    fs_on = False
    if viewer is not None:
        try:
            viewer.FullScreen = True
            fs_on = True
        except Exception:
            pass
    time.sleep(0.9)
    if fs_on and viewer is not None:
        # 全屏后视口变大, 重新拟合(最终仍按零件包围盒裁剪居中, 构图不受影响)
        try:
            viewer.Reframe()
        except Exception:
            pass
        _refresh()
        time.sleep(0.5)

    def _tree_visible(img_path):
        """检测结构树是否可见: 树首行(文档名)位于左上极角, 该区出现近黑像素即视为树可见。
        零件极少占据该极角, 故对零件无感。返回 True(可见)/False(已隐藏)/None(检测失败)。"""
        try:
            with Image.open(img_path) as _im:
                g = _im.convert("L")
                W, H = g.size
                band = g.crop((int(W * 0.02), 0, int(W * 0.22), int(H * 0.05)))
                px = list(band.getdata())
            return sum(1 for v in px if v < 120) >= 10
        except Exception:
            return None

    if viewer is not None:
        try:
            viewer.CaptureToFile(2, raw_path)
        except Exception:
            pass
    tv = _tree_visible(raw_path)
    if tv is True:
        # 兜底: FullScreen 未生效时再试 F3(多数情况下无效, 仅作最后手段)
        _hide_tree(catia, hwnd)
        time.sleep(0.5)
        log("  [视图] FullScreen 未隐藏结构树, 已尝试 F3 兜底")
    else:
        log("  [视图] 结构树已隐藏(FullScreen={})".format(fs_on))

    # 2) 实体着色: RenderingMode=6 (Shading with Edges, CATIA 默认风格) 或 5 (纯着色, 无实体边线)。
    #    实测本机 CATIA: 0~4 是线框/轮廓(完全没有表面色), 5/6 才是实体着色; 6 多一层实体边线。
    #    实体件用 6 最清晰; "曲面片拼接"的零件用 6 会把所有曲面片边界画出来、看着像网格,
    #    此时改用 5 得到干净的表面颜色渲染 —— 见下面的"边线主导"自动判定。
    n_frames = max(2, int(frames))

    def _grab_best(mode, n):
        """以指定渲染模式连抓 n 帧(每帧前强制重绘), 返回 _image_score 最高的 PIL 图 / None。
        帧覆写同一路径(raw_path), 不做批量临时文件删除。"""
        best, best_sc = None, -1.0
        for _ in range(max(2, int(n))):
            _set_mode(mode)
            _refresh()
            time.sleep(0.5)
            raw = raw_path
            got1 = False
            if viewer is not None:
                try:
                    viewer.CaptureToFile(2, raw)
                    if os.path.exists(raw) and os.path.getsize(raw) > 0:
                        got1 = True
                except Exception:
                    got1 = False
            if not got1:
                # 兜底: 屏幕抓取整个窗口
                try:
                    from PIL import ImageGrab
                    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
                    ImageGrab.grab(bbox=(left, top, right, bottom)).save(raw)
                    got1 = True
                except Exception:
                    got1 = capture_window(hwnd, raw)
            if got1 and os.path.exists(raw) and os.path.getsize(raw) > 0:
                try:
                    sc = _image_score(raw)
                except Exception:
                    sc = 0.0
                if sc > best_sc:
                    best_sc = sc
                    try:
                        with Image.open(raw) as _im:
                            _im.load()
                            best = _im.convert("RGB").copy()
                    except Exception:
                        best = None
        return best

    def _dark_nonwhite(im):
        """抽样统计 (暗像素数, 非白像素数)。"""
        try:
            g = im.convert("L")
            W, H = g.size
            px = g.load()
            dk = nw = 0
            for y in range(0, H, 2):
                for x in range(0, W, 2):
                    v = px[x, y]
                    if v < 120:
                        dk += 1
                    if v < 240:
                        nw += 1
            return dk, nw
        except Exception:
            return 0, 0

    # 4) 多帧捕获择优 (mode 6: 着色+实体边线)
    best_im = _grab_best(6, n_frames)
    try:
        _rm = viewer.RenderingMode if viewer is not None else "?"
    except Exception:
        _rm = "?"
    log("  [视图] RenderingMode(实)=%s (5/6=实体着色, 6 含实体边线)" % _rm)

    # 5) 后处理: 连通域抠出零件(去结构树/罗盘/坐标轴/背景) -> 裁剪居中 -> 800x600
    final_im, final_box = None, None
    if best_im is not None:
        try:
            final_im, final_box = _isolate_part(best_im)
        except Exception:
            final_im, final_box = None, None

    def _interior_edge_ratio(iso6, iso5):
        """零件"内部"(把零件掩码腐蚀掉约 7px 边界后的区域)中, mode6 相对 mode5 新增的暗像素占比。
        实体件的内部多为平坦面 -> 边线很少, 占比低;
        曲面片拼接件整套曲面片边界铺满内部 -> 占比高。返回比值或 None。"""
        try:
            g6 = iso6.convert("L")
            g5 = iso5.convert("L")
            if g5.size != g6.size:
                g5 = g5.resize(g6.size)
            mask = g6.point(lambda v: 255 if v < 240 else 0)          # 零件掩码
            interior = mask.filter(ImageFilter.MinFilter(15))         # 腐蚀 -> 内部
            dark6 = g6.point(lambda v: 255 if v < 120 else 0)
            dark5 = g5.point(lambda v: 255 if v < 120 else 0)
            area = interior.histogram()[255]
            if area <= 0:
                return None
            n6 = ImageChops.multiply(interior, dark6).histogram()[255]
            n5 = ImageChops.multiply(interior, dark5).histogram()[255]
            return (n6 - n5) / float(area)
        except Exception:
            return None

    # 5.1) "边线主导"判定: mode6 中暗像素占零件面积偏高时, 再抓一组 mode5(纯着色无边线),
    #      用"内部边线占比"区分两类情况 —— 曲面片拼接件(边线铺满内部)改用 mode5 得到干净的
    #      表面颜色渲染; 实体件(边线只在轮廓/孔洞, 内部干净)保留 mode6。深色实体件因 mode5
    #      下表面本身也是暗的, 差值≈0, 同样正确保留 mode6。
    #      阈值 0.09 由实测标定: 实体件 0.003~0.056, 曲面拼接件 0.125(两侧均有余量)。
    if final_im is not None:
        try:
            d6, nw6 = _dark_nonwhite(final_im)
            if nw6 > 0 and (d6 / float(nw6)) > 0.08:
                im5 = _grab_best(5, n_frames)
                if im5 is not None:
                    iso5, box5 = _isolate_part(im5)
                    er = _interior_edge_ratio(final_im, iso5)
                    if er is not None:
                        log("  [视图] 内部边线占比=%.3f (高=曲面拼接件)" % er)
                    if er is not None and er > 0.09:
                        final_im, final_box = iso5, box5
                        log("  [视图] 边线主导(曲面件) -> 改用 RenderingMode=5 纯着色")
        except Exception:
            pass

    if final_im is not None:
        try:
            out = _fit_center_on_white(final_im, final_box, PART_SHOT_SIZE)  # 裁剪居中缩放
            out.save(tmp)
        except Exception:
            try:
                _to_target_size(raw_path, tmp)
            except Exception:
                pass
        got = os.path.exists(tmp) and os.path.getsize(tmp) > 0
    else:
        got = False
    # 临时帧文件(raw)保留在磁盘不删除: 覆盖复用同一路径, 避免触发批量删除保护。
    # 退出全屏, 恢复常规界面(结构树/工具条) —— 不改变用户看到的 CATIA 界面
    if fs_on and viewer is not None:
        try:
            viewer.FullScreen = False
        except Exception:
            pass
        time.sleep(0.3)
    # 恢复窗口非置顶状态
    try:
        win32gui.SetWindowPos(hwnd, win32con.HWND_NOTOPMOST, 40, 40,
                              PART_SHOT_SIZE[0], PART_SHOT_SIZE[1],
                              win32con.SWP_NOMOVE | win32con.SWP_NOSIZE)
    except Exception:
        pass
    try:
        if sel is not None:
            sel.Clear()
    except Exception:
        pass
    return got


def capture_part(catia, doc, out_dir, base):
    """对整个零件截图: 单独窗口显示零件 -> 零件居中(FitAllIn) -> 隐藏结构树
    -> 白色背景 -> 截图(优先 CATIA 内置 CaptureToFile, 空白时回退屏幕抓取)。
    截图用于 Excel 测量结果行的"截图"列(与零件数据同一行)。返回图片路径或 None。"""
    stem = os.path.splitext(base)[0]
    path = os.path.join(out_dir, "{}__整机.png".format(sanitize_fname(stem)))
    if (not FORCE_CAPTURE) and os.path.exists(path) and os.path.getsize(path) > 0 and _is_target_size(path) and _image_has_content(path):
        return path  # 已按新规格截图且非空白, 跳过(支持断点续跑); --force 时忽略缓存
    try:
        os.makedirs(out_dir, exist_ok=True)
    except Exception:
        pass
    tmp = path + ".tmp.png"
    try:
        hwnd = find_catia_window(base)
        if hwnd is None:
            log("  [零件截图失败] 未找到 CATIA 窗口: " + base)
            return None
        for attempt in range(3):
            if _capture_view(catia, doc, hwnd, tmp, base, fit_center=True):
                _to_target_size(tmp, path)
                return path
            log("  [零件截图] 第 {} 次尝试失败, 重试...".format(attempt + 1))
            time.sleep(1.0)  # 等待视口渲染完成后再试
        log("  [零件截图失败] 多次尝试仍无法截取有效视图: " + base)
        return None
    except Exception as e:
        log("  [零件截图失败] {}: {}".format(base, repr(e)[:120]))
        return None


def capture_annotation(catia, doc, annotation, out_dir, base, ann_name):
    """选中单个标注后按统一视图截图(窗口 800x600、零件居中、隐藏结构树、白色背景),
    以"标注名__零件名.png"保存。返回图片路径或 None。"""
    stem = os.path.splitext(base)[0]
    path = os.path.join(out_dir, "{}__{}.png".format(sanitize_fname(ann_name), sanitize_fname(stem)))
    if (not FORCE_CAPTURE) and os.path.exists(path) and os.path.getsize(path) > 0 and _is_target_size(path) and _image_has_content(path):
        return path  # 已按新规格截图且非空白, 跳过(支持断点续跑); --force 时忽略缓存
    tmp = path + ".tmp.png"
    try:
        hwnd = find_catia_window(base)
        if hwnd is None:
            log("    [截图失败] 未找到 CATIA 窗口: " + base)
            return None
        for attempt in range(2):
            if _capture_view(catia, doc, hwnd, tmp, base, fit_center=True, select=[annotation], frames=1):
                _to_target_size(tmp, path)
                return path
            log("    [截图] 第 {} 次尝试失败, 重试: {}".format(attempt + 1, ann_name))
            time.sleep(0.8)
        log("    [截图失败] 多次尝试仍无法截取有效视图: " + ann_name)
        return None
    except Exception as e:
        log("    [截图失败] {}: {}".format(ann_name, repr(e)[:120]))
        return None


def capture_all_annotations(catia, doc, out_dir, base):
    """选中零件 AnnotationSets 下所有标注后按统一视图截图(窗口 800x600、
    零件居中、隐藏结构树、白色背景), 以"零件名__全部标注.png"保存。
    返回图片路径或 None。"""
    path = os.path.join(out_dir, "{}__全部标注.png".format(sanitize_fname(os.path.splitext(base)[0])))
    if (not FORCE_CAPTURE) and os.path.exists(path) and os.path.getsize(path) > 0 and _is_target_size(path) and _image_has_content(path):
        return path  # 已按新规格截图且非空白, 跳过(支持断点续跑); --force 时忽略缓存
    tmp = path + ".tmp.png"
    try:
        hwnd = find_catia_window(base)
        if hwnd is None:
            log("  [全部标注截图失败] 未找到 CATIA 窗口: " + base)
            return None
        # 收集所有标注对象
        anns_all = []
        sets = doc.Part.AnnotationSets
        for si in range(1, sets.Count + 1):
            try:
                anns = sets.Item(si).Annotations
            except Exception:
                continue
            for ai in range(1, anns.Count + 1):
                try:
                    anns_all.append(anns.Item(ai))
                except Exception:
                    continue
        if not anns_all:
            log("  [全部标注截图] 无标注可选中: " + base)
            return None
        for attempt in range(2):
            if _capture_view(catia, doc, hwnd, tmp, base, fit_center=True, select=anns_all):
                _to_target_size(tmp, path)
                return path
            log("  [全部标注截图] 第 {} 次尝试失败, 重试...".format(attempt + 1))
            time.sleep(0.8)
        log("  [全部标注截图失败] 多次尝试仍无法截取有效视图: " + base)
        return None
    except Exception as e:
        log("  [全部标注截图失败] {}: {}".format(base, repr(e)[:120]))
        return None


def shots_for_part(catia, part_path, shots_dir, annotations, part_shots_dir=None):
    """重新打开零件, 依次截: 整机图、全部标注(全选)图、每个三维标注图。
    返回 (整机截图路径或 None, 全部标注截图路径或 None, [(标注名, 图片路径), ...])"""
    doc = None
    results = []
    part_shot = None
    all_shot = None
    base = os.path.basename(part_path)
    try:
        doc = _open_document(catia, part_path)
        time.sleep(2.0)  # 等待 3D 视图渲染完成, 否则截图可能抓到未重绘的旧画面
        if part_shots_dir:
            part_shot = capture_part(catia, doc, part_shots_dir, base)
            # 选中所有标注的整体截图(800x600, 白背景, 无结构树)
            all_shot = capture_all_annotations(catia, doc, part_shots_dir, base)
        part = doc.Part
        try:
            sets = part.AnnotationSets
            for si in range(1, sets.Count + 1):
                try:
                    anns = sets.Item(si).Annotations
                except Exception:
                    continue
                for ai in range(1, anns.Count + 1):
                    try:
                        a = anns.Item(ai)
                    except Exception:
                        continue
                    ann_name = ""
                    try:
                        ann_name = str(a.Name)
                    except Exception:
                        pass
                    if not ann_name:
                        ann_name = "标注_{}_{}".format(si, ai)
                    # 仅截图中存在的标注
                    if not any(n == ann_name for _, n, _ in annotations):
                        continue
                    log("    截图: " + ann_name)
                    p = capture_annotation(catia, doc, a, shots_dir, os.path.basename(part_path), ann_name)
                    if p:
                        results.append((ann_name, p))
        except Exception:
            pass
    finally:
        if doc is not None:
            try:
                doc.Close()
            except Exception:
                pass
    return part_shot, all_shot, results


# ============================ MBD 导出(源自 MBD.txt VBA 移植) ============================

# MBD 指定参数的"所属集合"关键字(与 VBA targetSets 一致), 可在 GUI 中勾选
MBD_PARAM_SETS = [
    "Part Notes:",
    "Standard Notes:",
    "Annotation Notes:",
    "Material Description:",
    "ECCN",
    "Approval Status",
]


def collect_mbd_params(part, target_sets=None):
    """提取指定集合下的参数(VBA 功能块 1 的 Python 移植)。
    遍历全部参数, 名称中包含任一目标集合关键字即提取; 值读取优先
    ValueAsString() 方法(含单位), 失败回退 .Value 属性。
    返回 [(所属集合, 显示名称, 值字符串), ...]"""
    if target_sets is None:
        target_sets = MBD_PARAM_SETS
    out = []
    try:
        params = part.Parameters
    except Exception:
        return out
    for i in range(1, params.Count + 1):
        try:
            p = params.Item(i)
            nm = str(p.Name)
        except Exception:
            continue
        hit = ""
        for t in target_sets:
            if t.lower() in nm.lower():
                hit = t
                break
        if not hit:
            continue
        display = nm.split("\\")[-1] if "\\" in nm else nm
        val = "N/A"
        try:
            val = str(p.ValueAsString())
        except Exception:
            try:
                val = str(p.Value)
            except Exception:
                pass
        out.append((hit, display, val))
    return out


def capture_capture_view(catia, o_capture, index, out_dir, wait_sec=1.0):
    """激活单个 Capture 视图并截图(VBA 功能块 2 的 Python 移植)。
    o_capture.DisplayCapture -> 等待渲染 -> ActiveViewer.CaptureToFile 截 JPG。
    返回图片路径; 失败或空图返回 None。"""
    try:
        o_capture.DisplayCapture()
    except Exception:
        return None
    try:
        catia.RefreshDisplay = True
    except Exception:
        pass
    time.sleep(wait_sec)
    if not out_dir:
        out_dir = tempfile.gettempdir()
    os.makedirs(out_dir, exist_ok=True)
    tmp = os.path.join(out_dir, "CATIA_Capture_{}.jpg".format(index))
    try:
        viewer = catia.ActiveWindow.ActiveViewer
        viewer.CaptureToFile(5, tmp)  # catCaptureFormatJPEG = 5(与 VBA 一致)
    except Exception:
        return None
    if not os.path.exists(tmp) or os.path.getsize(tmp) == 0:
        return None
    return tmp


# ============================ Excel 输出 ============================

def sanitize_sheet_name(name, used):
    """生成合法的 Excel sheet 名(<=31 字符, 去非法字符), 冲突时追加序号"""
    s = re.sub(r'[\\/:*?\[\]]', "_", str(name)).strip()[:31] or "零件"
    base = s
    i = 2
    while s in used:
        s = base[:28] + "_" + str(i)
        i += 1
    used.add(s)
    return s


def part_number_of(d):
    """取零件号: 优先用户参数"零件编号", 否则取文件名第一个下划线前的字段"""
    for name, val in d.get("user_params") or []:
        if name == "零件编号" and str(val).strip():
            return str(val).strip()
    return d["file"].split("_")[0]

def _ensure_shot_size(path, size=PART_SHOT_SIZE):
    """把截图统一为指定尺寸(默认 800x600, 等比缩放后白底补齐), 原地覆盖保存。"""
    try:
        with Image.open(path) as im:
            if im.size == size:
                return
            im.load()
        scale = min(size[0] / im.size[0], size[1] / im.size[1])
        nw, nh = max(1, int(round(im.size[0] * scale))), max(1, int(round(im.size[1] * scale)))
        im2 = im.resize((nw, nh), Image.LANCZOS)
        canvas = Image.new("RGB", size, (255, 255, 255))
        canvas.paste(im2, ((size[0] - nw) // 2, (size[1] - nh) // 2))
        canvas.save(path)
    except Exception:
        pass


def _insert_shot(ws, row, col, path, row_height=SHOT_ROW_HEIGHT):
    """把截图插入 Excel, 并使其显示高度与所在行行高一致(默认 30 点), 等比缩放。
    xlsxwriter 图片按 96DPI 解释, Excel 行高按 72DPI, 故像素高度 pt = _ih*0.75。"""
    try:
        with Image.open(path) as _im:
            _iw, _ih = _im.size
        if _ih <= 0:
            return False
        _s = row_height / (_ih * 0.75)   # 等比缩放到"显示高度=行高"
        ws.insert_image(row, col, path, {"x_scale": _s, "y_scale": _s})
        ws.set_row(row, row_height)
        return True
    except Exception:
        return False


def write_excel(out_path, results, shots_map=None, sheets=None):
    """生成 Excel:
      测量结果  每零件一行, 含整机截图列(截图文件 800x600, 表格内显示高度=行高)
      用户参数  每个零件的用户自定义参数列表(零件文件/参数名/参数值)
      MBD参数  几何图形集参数 + MBD 指定参数合并去重(零件文件/所属集合/参数名/参数值)
      参数汇总  列头=参数名, 行=零件, 单元格=参数值
      三维标注  单一 sheet 页, 每零件含"全部标注"整体截图与逐个标注截图
      MBD捕获截图  按件号分页, 每页含捕获名称与 800x600 截图
    sheets: 要输出的页面列表, 缺省输出全部;
            可选: 测量结果/用户参数/参数汇总/三维标注/MBD参数/MBD捕获截图"""
    if sheets is None:
        sheets = ["测量结果", "用户参数", "参数汇总", "三维标注", "MBD参数", "MBD捕获截图"]
    shots_map = shots_map or {}
    wb = xlsxwriter.Workbook(out_path)
    header_fmt = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1, "align": "center"})
    cell = wb.add_format({"border": 1})

    # ---- 测量结果 ----
    if "测量结果" in sheets:
        ws = wb.add_worksheet("测量结果")
        headers = ["文件", "类型", "体积(cm3)", "表面积(cm2)", "密度(g/cm3)", "密度来源",
                   "质量(kg)", "重量(N)", "包络体长(mm)", "包络体宽(mm)", "包络体高(mm)", "备注", "截图"]
        for c, h in enumerate(headers):
            ws.write(0, c, h, header_fmt)
        r = 1
        for d in results:
            box = d.get("box") or []
            vals = [
                d["file"],
                "零件" if d["kind"] == "Part" else "装配体",
                round(d["vol_m3"] * 1e6, 3),
                round(d["area_m2"] * 1e4, 3) if d["kind"] == "Part" else "",
                round(d["density_g_cm3"], 3),
                d["density_src"],
                round(d["mass_kg"], 4),
                round(d["weight_n"], 3) if d["kind"] == "Part" else "",
                round(box[0], 3) if len(box) > 0 else "",
                round(box[1], 3) if len(box) > 1 else "",
                round(box[2], 3) if len(box) > 2 else "",
                d["error"] or d["note"],
                "",
            ]
            for c, v in enumerate(vals):
                ws.write(r, c, v, cell)
            ws.set_row(r, SHOT_ROW_HEIGHT)   # 每行行高统一 = SHOT_ROW_HEIGHT(60pt)
            # 零件整机截图插入"截图"列, 显示高度与行高一致(60pt), 等比缩放
            shot = d.get("part_shot") or ""
            if shot and os.path.exists(shot):
                try:
                    _insert_shot(ws, r, len(headers) - 1, shot)
                except Exception:
                    pass
            r += 1
        ws.set_column(0, 0, 42)
        ws.set_column(1, 1, 8)
        ws.set_column(2, 10, 12)
        ws.set_column(11, 11, 30)
        ws.set_column(12, 12, 17)

    # ---- 用户参数(每个零件的参数列表) ----
    if "用户参数" in sheets:
        ws_up = wb.add_worksheet("用户参数")
        h_up = ["零件文件", "参数名", "参数值"]
        for c, h in enumerate(h_up):
            ws_up.write(0, c, h, header_fmt)
        r_up = 1
        any_up = False
        for d in results:
            if d["kind"] != "Part" or not d.get("user_params"):
                continue
            any_up = True
            for name, val in d["user_params"]:
                ws_up.write_row(r_up, 0, [d["file"], name, val], cell)
                r_up += 1
        if not any_up:
            ws_up.write(1, 0, "(未检测到用户参数)", cell)
        ws_up.set_column(0, 0, 42)
        ws_up.set_column(1, 1, 32)
        ws_up.set_column(2, 2, 24)

    # ---- 几何图形集参数(挂在"几何图形集/图形数据集合"节点上的参数与参数值) ----
    if "几何图形集参数" in sheets:
        ws_g = wb.add_worksheet("几何图形集参数")
        h_g = ["零件文件", "几何图形集", "参数名", "参数值"]
        for c, h in enumerate(h_g):
            ws_g.write(0, c, h, header_fmt)
        r_g = 1
        any_g = False
        for d in results:
            if d["kind"] != "Part" or not d.get("geomset_params"):
                continue
            any_g = True
            for name, val in d["geomset_params"]:
                # name = '几何图形集名\参数名'
                owner, _, pname = name.rpartition("\\")
                ws_g.write_row(r_g, 0, [d["file"], owner, pname, val], cell)
                r_g += 1
        if not any_g:
            ws_g.write(1, 0, "(未检测到几何图形集参数)", cell)
        ws_g.set_column(0, 0, 42)
        ws_g.set_column(1, 1, 30)
        ws_g.set_column(2, 2, 30)
        ws_g.set_column(3, 3, 80)

    # ---- 参数汇总(列头=参数名, 行=零件, 单元格=参数值) ----
    if "参数汇总" in sheets:
        ws_ps = wb.add_worksheet("参数汇总")
        param_names = []
        seen = set()
        for d in results:
            if d["kind"] != "Part":
                continue
            for name, _ in d.get("user_params") or []:
                if name not in seen:
                    seen.add(name)
                    param_names.append(name)
        ws_ps.write(0, 0, "零件文件", header_fmt)
        for c, name in enumerate(param_names, start=1):
            ws_ps.write(0, c, name, header_fmt)
        r_ps = 1
        any_ps = False
        for d in results:
            if d["kind"] != "Part":
                continue
            any_ps = True
            pvals = dict(d.get("user_params") or [])
            ws_ps.write(r_ps, 0, d["file"], cell)
            for c, name in enumerate(param_names, start=1):
                ws_ps.write(r_ps, c, pvals.get(name, ""), cell)
            r_ps += 1
        if not any_ps:
            ws_ps.write(1, 0, "(未检测到用户参数)", cell)
        ws_ps.set_column(0, 0, 42)
        ws_ps.set_column(1, max(1, len(param_names)), 20)

    # ---- 三维标注(单一 sheet 页, 不含截图列) ----
    if "三维标注" in sheets:
        ws2 = wb.add_worksheet("三维标注")
        h2 = ["零件文件", "标注集", "标注名称", "标注文本"]
        for c, h in enumerate(h2):
            ws2.write(0, c, h, header_fmt)
        rx = 1
        any_ann = False
        for d in results:
            if d["kind"] != "Part" or not d.get("annotations"):
                continue
            any_ann = True
            # 首行: 全部标注(选中所有)
            ws2.write_row(rx, 0, [d["file"], "—", "全部标注(选中所有)", "—"], cell)
            rx += 1
            # 逐个标注行
            for set_name, ann_name, ann_text in d["annotations"]:
                ws2.write_row(rx, 0, [d["file"], set_name, ann_name, ann_text], cell)
                rx += 1
        if not any_ann:
            ws2.write(1, 0, "(未检测到三维标注)", cell)
        ws2.set_column(0, 0, 42)
        ws2.set_column(1, 1, 16)
        ws2.set_column(2, 2, 36)
        ws2.set_column(3, 3, 40)
    # ---- MBD 参数(合并"几何图形集参数"与"MBD 指定参数", 按行内容去重) ----
    if "MBD参数" in sheets:
        ws_mb = wb.add_worksheet("MBD参数")
        h_mb = ["零件文件", "所属集合/几何图形集", "参数名称", "参数值"]
        for c, h in enumerate(h_mb):
            ws_mb.write(0, c, h, header_fmt)
        r_mb = 1
        any_mb = False
        seen_rows = set()
        for d in results:
            if d["kind"] != "Part":
                continue
            # 行格式统一为 (file, 所属集合或几何图形集, 参数名, 参数值)
            rows = []
            for hit, display, val in d.get("mbd_params") or []:
                rows.append((d["file"], hit, display, val))
            for name, val in d.get("geomset_params") or []:
                # name = '几何图形集名\参数名'
                owner, _, pname = name.rpartition("\\")
                rows.append((d["file"], owner, pname, val))
            for row in rows:
                if row in seen_rows:
                    continue
                seen_rows.add(row)
                any_mb = True
                ws_mb.write_row(r_mb, 0, list(row), cell)
                r_mb += 1
        if not any_mb:
            ws_mb.write(1, 0, "(未检测到参数)", cell)
        ws_mb.set_column(0, 0, 42)
        ws_mb.set_column(1, 1, 30)
        ws_mb.set_column(2, 2, 32)
        ws_mb.set_column(3, 3, 60)

    # ---- MBD 捕获截图(按件号分页, 截图后追加该零件的 MBD 参数数据) ----
    if "MBD捕获截图" in sheets:
        mbd_used = set()
        mbd_any = False
        for d in results:
            caps_map = d.get("mbd_capture_shots") or {}
            if d["kind"] != "Part":
                continue
            # 该零件的 MBD 参数行(与"MBD参数"页同源: MBD 指定参数 + 几何图形集参数, 去重)
            param_rows = []
            for hit, display, val in d.get("mbd_params") or []:
                param_rows.append((hit, display, val))
            for name, val in d.get("geomset_params") or []:
                owner, _, pname = name.rpartition("\\")
                param_rows.append((owner, pname, val))
            if not caps_map and not param_rows:
                continue
            mbd_any = True
            sname = sanitize_sheet_name(part_number_of(d), mbd_used)
            ws_mc = wb.add_worksheet(sname)
            h_mc = ["捕获名称", "截图文件", "截图"]
            for c, h in enumerate(h_mc):
                ws_mc.write(0, c, h, header_fmt)
            r_mc = 1
            for cap_name, img in caps_map.items():
                ws_mc.write_row(r_mc, 0, [cap_name, img], cell)
                if img and os.path.exists(img):
                    try:
                        _insert_shot(ws_mc, r_mc, 2, img, row_height=MBD_SHOT_ROW_HEIGHT)
                    except Exception:
                        pass
                r_mc += 1
            # 截图后追加该零件的 MBD 参数数据
            if param_rows:
                r_mc += 1  # 空一行分隔
                for c, h in enumerate(["所属集合/几何图形集", "参数名称", "参数值"]):
                    ws_mc.write(r_mc, c, h, header_fmt)
                r_mc += 1
                for hit, pname, val in param_rows:
                    ws_mc.write_row(r_mc, 0, [hit, pname, val], cell)
                    r_mc += 1
            ws_mc.set_column(0, 0, 32)
            ws_mc.set_column(1, 1, 46)
            ws_mc.set_column(2, 2, 80)
        if not mbd_any:
            ws_mc = wb.add_worksheet("MBD捕获截图")
            ws_mc.write(1, 0, "(未检测到 Capture 或未勾选该提取项)", cell)
    wb.close()


# ============================ 主流程 ============================

def _is_rpc_error(msg):
    """判断错误信息是否为 CATIA 进程崩溃/断连类 COM 错误(RPC 服务器不可用等)。
    CATIA 中途崩溃后, 后续所有 COM 调用都会立刻抛这类错误; 必须重启 CATIA 才能恢复。"""
    s = str(msg)
    if "RPC" in s or "0x800706BA" in s or "0x80010108" in s:
        return True
    try:
        code = int(getattr(msg, "hresult", -1))
    except Exception:
        code = -1
    return code in (-2147023174, -2147417848, -2147418111)  # RPC_S_SERVER_UNAVAILABLE / DISCONNECTED / CALL_CANCELED


def run_measure(args, catia, d, files, out_txt, json_path, opts=None):
    header = [
        "=" * 60,
        "CATIA 测量结果  时间: {}".format(time.strftime("%Y-%m-%d %H:%M:%S")),
        "目录: {}".format(d),
        "密度: {} g/cm3 (默认; 零件已赋材质时自动采用材质密度)".format(args.density),
        "包络体: 与坐标轴对齐的 AABB, 单位 mm (优先 GetBoundingBox, 已排除远置辅助体; 异常值仍会标记疑似)",
        "=" * 60,
        "",
    ]
    results = []
    ok_count = 0
    with open(out_txt, "w", encoding="utf-8") as f:
        f.write("\n".join(header))
        for fn in files:
            ok, text, data = measure_one(catia, os.path.join(d, fn), args.density, opts)
            if not ok and _is_rpc_error(data.get("error", "")):
                # CATIA 进程崩溃(RPC 断连): 不再让剩余文件逐个快速失败,
                # 自动重启 CATIA 并重试当前文件一次。
                log("  [恢复] CATIA 连接断开, 自动重启 CATIA 后重试: " + fn)
                try:
                    catia = connect_catia()
                except Exception as ce:
                    log("  [恢复] CATIA 重启失败, 继续处理剩余文件: " + repr(ce)[:120])
                ok, text, data = measure_one(catia, os.path.join(d, fn), args.density, opts)
            results.append(data)
            if ok:
                ok_count += 1
            f.write(text)
            f.flush()
            log(text)
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(results, f, ensure_ascii=False, indent=1)
    log("=" * 60)
    log("完成: {}/{} 个文件测量成功".format(ok_count, len(files)))
    log("文本报告: " + out_txt)
    return results


def run_shots(args, catia, d, json_path, shots_dir, out_xlsx):
    """读取测量结果, 逐个零件截整机图 + 三维标注截图, 重新生成 Excel"""
    with open(json_path, "r", encoding="utf-8") as f:
        results = json.load(f)
    # 清理 CATIA 中遗留文档, 避免被中断的进程留下异常状态阻塞后续操作
    for i in range(catia.Documents.Count, 0, -1):
        try:
            catia.Documents.Item(i).Close()
        except Exception:
            pass
    os.makedirs(shots_dir, exist_ok=True)
    part_shots_dir = os.path.join(d, "零件截图")
    os.makedirs(part_shots_dir, exist_ok=True)
    shots_map = {}
    for dta in results:
        if dta.get("kind") != "Part":
            continue
        part_path = os.path.join(d, dta["file"])
        if not os.path.exists(part_path):
            continue
        log("截图: " + dta["file"])
        part_shot, all_shot, ann_results = shots_for_part(
            catia, part_path, shots_dir, dta.get("annotations") or [], part_shots_dir)
        if part_shot:
            dta["part_shot"] = part_shot
        if all_shot:
            dta["all_shot"] = all_shot
        for ann_name, img in ann_results:
            shots_map[(dta["file"], ann_name)] = img
    write_excel(out_xlsx, results, shots_map)
    log("零件整机截图: {} 张, 三维标注截图: {} 张".format(
        sum(1 for x in results if x.get("part_shot")), len(shots_map)))
    log("Excel 已更新: " + out_xlsx)


def main():
    ap = argparse.ArgumentParser(description="控制 CATIA 测量零件/装配体并导出 Excel")
    ap.add_argument("--dir", default="test", help="要测量的目录 (默认 test)")
    ap.add_argument("--density", type=float, default=DEFAULT_DENSITY, help="密度 g/cm3 (默认 1.0)")
    ap.add_argument("--out", default=None, help="文本报告路径 (默认 <dir>/测量结果.txt)")
    ap.add_argument("--excel", default=None, help="Excel 路径 (默认 <dir>/测量结果.xlsx)")
    ap.add_argument("--json", default=None, help="中间数据路径 (默认 <dir>/results.json)")
    ap.add_argument("--shots", action="store_true", help="三维标注截图并更新 Excel (需先运行主程序)")
    ap.add_argument("--no-part-shot", action="store_true",
                    help="不生成整机截图(仅出几何/参数数据, 加快批量测量; 默认会截整机图并写入行)")
    ap.add_argument("--gui", action="store_true", help="打开图形界面(复选框选择要提取的内容)")
    ap.add_argument("--force", action="store_true",
                    help="强制重新生成所有截图(忽略已存在的截图缓存, 用于修复着色/尺寸后重跑)")
    args = ap.parse_args()
    global FORCE_CAPTURE
    FORCE_CAPTURE = bool(args.force)

    if args.gui:
        try:
            import measure_gui
        except Exception as e:
            log("无法启动界面: {}".format(repr(e)[:200]))
            sys.exit(1)
        measure_gui.run_gui()
        return

    d = os.path.abspath(args.dir)
    if not os.path.isdir(d):
        log("目录不存在: " + d)
        sys.exit(1)

    out_txt = args.out or os.path.join(d, "测量结果.txt")
    out_xlsx = args.excel or os.path.join(d, "测量结果.xlsx")
    json_path = args.json or os.path.join(d, "results.json")
    shots_dir = os.path.join(d, "三维标注截图")

    pythoncom.CoInitialize()
    catia = connect_catia()

    if args.shots:
        run_shots(args, catia, d, json_path, shots_dir, out_xlsx)
    else:
        files = sorted(f for f in os.listdir(d) if f.lower().endswith((".catpart", ".catproduct")))
        if not files:
            log("目录中没有 CATPart/CATProduct 文件: " + d)
            sys.exit(1)
        opts = {"part_shot": not args.no_part_shot}
        results = run_measure(args, catia, d, files, out_txt, json_path, opts)
        write_excel(out_xlsx, results, {})
        log("Excel: " + out_xlsx)
        if any(dta.get("kind") == "Part" and dta.get("annotations") for dta in results):
            log("提示: 运行 'python measure_catia.py --shots' 可生成三维标注截图并写入 Excel")

    try:
        catia.DisplayFileAlerts = True
    except Exception:
        pass


if __name__ == "__main__":
    main()
