import streamlit as st
import json
import os
import base64
import tempfile
import time
import zipfile
import shutil
import re
from pathlib import Path
import anthropic

# ===================== 页面配置 =====================
st.set_page_config(
    page_title="房地产评估报告生成系统",
    page_icon="🏠",
    layout="wide"
)

st.markdown("""
<style>
    .main-title {
        text-align: center;
        color: #2c3e50;
        font-size: 2rem;
        font-weight: bold;
        padding: 1rem 0;
        border-bottom: 3px solid #3498db;
        margin-bottom: 2rem;
    }
    .section-card {
        background: #f8f9fa;
        border-left: 4px solid #3498db;
        padding: 1rem 1.5rem;
        border-radius: 0 8px 8px 0;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)


# ===================== 模板占位符定义 =====================
# 与 template_v2.docx 中的占位符完全对应
TEMPLATE_PATH = Path(__file__).parent / "template_v2.docx"

# 表单分组展示
FIELD_GROUPS = {
    "📋 基本信息": [
        "权属人", "房产地址", "委托人", "委托书文号",
        "报告编号", "报告序号", "价值时点", "报告日期",
        "查勘日期", "签名日期", "作业期",
    ],
    "🏠 房产实物": [
        "建筑面积", "土地面积", "户型", "总层数",
        "所在楼层", "建成年份", "欠缴特约物业费", "欠缴物业费",
    ],
    "📜 权益状况": [
        "不动产权证号", "使用期限", "宗地号", "登记日期",
        "抵押权人", "抵押登记证明号", "债权数额",
        "抵押登记日期", "债务履行期限", "查封文号", "查封期限",
    ],
    "💰 估价结论": [
        "评估总价", "评估总价大写", "评估单价", "评估单价大写",
    ],
}

# 所有字段平铺列表（供提取prompt使用）
ALL_FIELDS = [f for fields in FIELD_GROUPS.values() for f in fields]


# ===================== 工具函数 =====================

def pdf_to_images_base64(pdf_path: str, scale: float = 1.2) -> list:
    """将PDF各页转换为base64图片，scale控制分辨率（越小越省流量）"""
    import fitz
    doc = fitz.open(pdf_path)
    images = []
    for page_num in range(len(doc)):
        page = doc.load_page(page_num)
        pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale))
        b64 = base64.standard_b64encode(pix.tobytes("jpeg", jpg_quality=75)).decode("utf-8")
        images.append({"page": page_num + 1, "base64": b64, "media_type": "image/jpeg"})
    doc.close()
    return images


def extract_info_from_pdf(pdf_path: str, api_key: str) -> dict:
    """用Claude Vision从扫描PDF中提取结构化房产信息，分批处理避免超时"""
    client = anthropic.Anthropic(api_key=api_key, base_url="https://api.302.ai")

    with st.spinner("📄 正在将PDF转换为图片..."):
        images = pdf_to_images_base64(pdf_path, scale=1.2)
    st.info(f"共转换 {len(images)} 页，开始AI识别...")

    json_template = json.dumps({f: "" for f in ALL_FIELDS}, ensure_ascii=False, indent=2)
    extract_prompt = f"""请仔细识别这份房地产估价PDF文档，以JSON格式返回以下字段（未提及填"未提及"）：
{json_template}
只返回JSON，不要任何其他文字。"""

    # 分批处理：每批最多3页，避免单次请求过大超时
    BATCH_SIZE = 3
    all_results = {}

    for batch_start in range(0, len(images), BATCH_SIZE):
        batch = images[batch_start: batch_start + BATCH_SIZE]
        batch_end = batch_start + len(batch)
        with st.spinner(f"🤖 正在识别第 {batch_start+1}-{batch_end} 页..."):
            content = []
            for img in batch:
                content.append({
                    "type": "image",
                    "source": {"type": "base64", "media_type": img["media_type"], "data": img["base64"]}
                })
            if batch_start == 0:
                content.append({"type": "text", "text": extract_prompt})
            else:
                already = json.dumps(all_results, ensure_ascii=False)
                content.append({"type": "text", "text": f"""这是文档后续页面，请补充提取之前未能获取的字段。
已提取到的信息：{already}
请从这些页面中提取仍为空或"未提及"的字段，返回JSON，只包含有新值的字段。"""})

            response = client.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=2000,
                messages=[{"role": "user", "content": content}]
            )
            raw = response.content[0].text.strip()
            if "```json" in raw:
                raw = raw.split("```json")[1].split("```")[0].strip()
            elif "```" in raw:
                raw = raw.split("```")[1].split("```")[0].strip()
            try:
                batch_result = json.loads(raw)
                for k, v in batch_result.items():
                    if v and v != "未提及" and (k not in all_results or not all_results[k] or all_results[k] == "未提及"):
                        all_results[k] = v
            except Exception:
                pass

    for f in ALL_FIELDS:
        if f not in all_results:
            all_results[f] = "未提及"

    return all_results


def search_surroundings(address: str, amap_key: str) -> dict:
    """高德地图周边搜索"""
    import requests
    result = {
        "坐标": None,
        "交通（地铁/公交）": [],
        "教育（学校/幼儿园）": [],
        "医疗（医院/诊所）": [],
        "商业（商场/超市）": [],
        "公园绿地": [],
        "搜索状态": "成功"
    }
    try:
        geo = requests.get(
            "https://restapi.amap.com/v3/geocode/geo",
            params={"address": address, "key": amap_key, "output": "json"},
            timeout=10
        ).json()
        if geo.get("status") != "1" or not geo.get("geocodes"):
            result["搜索状态"] = "地址解析失败"
            return result
        location = geo["geocodes"][0]["location"]
        result["坐标"] = location

        for type_str, key, radius in [
            ("交通设施服务",            "交通（地铁/公交）", 1000),
            ("中小学;高等院校;幼儿园",  "教育（学校/幼儿园）", 1000),
            ("综合医院;诊所;药店",      "医疗（医院/诊所）",  1500),
            ("购物服务;超级市场",       "商业（商场/超市）",  1000),
            ("公园广场;风景名胜",       "公园绿地",          1500),
        ]:
            resp = requests.get(
                "https://restapi.amap.com/v3/place/around",
                params={"location": location, "types": type_str, "radius": radius,
                        "key": amap_key, "output": "json", "offset": 5},
                timeout=10
            ).json()
            if resp.get("status") == "1" and resp.get("pois"):
                for poi in resp["pois"][:5]:
                    result[key].append(f"{poi.get('name','')}（约{poi.get('distance','')}米）")
    except Exception as e:
        result["搜索状态"] = f"搜索异常: {e}"
    return result


def generate_surrounding_description(info: dict, surroundings: dict, api_key: str) -> tuple:
    """用Claude生成区位描述两段，返回 (段落1, 段落2)"""
    client = anthropic.Anthropic(api_key=api_key, base_url="https://api.302.ai")
    prompt = f"""根据以下房产信息和周边配套数据，为房地产估价报告生成"区位状况描述与分析"内容。

房产信息：{json.dumps(info, ensure_ascii=False)}
周边配套：{json.dumps(surroundings, ensure_ascii=False)}

请生成两段内容，用 ---SPLIT--- 分隔：
第一段（约150字）：描述估价对象的具体区位，包括所处小区四至方位、周边住宅小区、基础设施、公共服务设施、交通、商业配套等。
第二段（约80字）：从整体区位状况做综合评价，包括居住氛围、人文环境、自然环境、未来趋势等。

直接输出正文，不要标题，两段之间用 ---SPLIT--- 分隔。"""

    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=800,
        messages=[{"role": "user", "content": prompt}]
    )
    text = response.content[0].text.strip()
    parts = text.split("---SPLIT---")
    para1 = parts[0].strip() if parts else text
    para2 = parts[1].strip() if len(parts) > 1 else ""
    return para1, para2


def fill_template(data: dict, output_path: str):
    """将数据填入 template_v2.docx 模板，替换所有 {{占位符}}"""
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(
            f"模板文件不存在：{TEMPLATE_PATH}\n"
            "请将 template_v2.docx 放在程序同目录下。"
        )
    shutil.copy(str(TEMPLATE_PATH), output_path)

    with zipfile.ZipFile(output_path, 'r') as z:
        xml_content = z.read('word/document.xml').decode('utf-8')

    for key, value in data.items():
        xml_content = xml_content.replace("{{" + key + "}}", str(value) if value else "")

    tmp_path = output_path + ".tmp"
    with zipfile.ZipFile(output_path, 'r') as zin:
        with zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == 'word/document.xml':
                    zout.writestr(item, xml_content.encode('utf-8'))
                else:
                    zout.writestr(item, zin.read(item.filename))
    os.replace(tmp_path, output_path)


def replace_image_in_docx(docx_path: str, image_placeholder: str,
                           new_image_bytes: bytes, image_ext: str = "jpeg"):
    """替换模板中指定标识（IMAGE_xxx）的图片"""
    with zipfile.ZipFile(docx_path, 'r') as z:
        rels_content = z.read('word/_rels/document.xml.rels').decode('utf-8')
        doc_content  = z.read('word/document.xml').decode('utf-8')

    # 找到含placeholder的blip对应的rId
    match = re.search(
        rf'r:embed="(rId\d+)"[^>]*w:comment="{image_placeholder}"', doc_content
    )
    if not match:
        return
    rid = match.group(1)

    # 找到rels中对应的文件名
    rels_match = re.search(rf'Id="{rid}"[^>]*Target="media/([^"]+)"', rels_content)
    if not rels_match:
        return
    old_filename = rels_match.group(1)
    new_filename = f"replaced_{image_placeholder.lower()}.{image_ext}"

    rels_content = rels_content.replace(
        f'media/{old_filename}', f'media/{new_filename}'
    )

    tmp_path = docx_path + ".imgtmp"
    with zipfile.ZipFile(docx_path, 'r') as zin:
        with zipfile.ZipFile(tmp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == 'word/_rels/document.xml.rels':
                    zout.writestr(item, rels_content.encode('utf-8'))
                elif item.filename == f'word/media/{old_filename}':
                    zout.writestr(f'word/media/{new_filename}', new_image_bytes)
                else:
                    zout.writestr(item, zin.read(item.filename))
    os.replace(tmp_path, docx_path)


# ===================== 主界面 =====================

st.markdown('<div class="main-title">🏠 房地产评估报告生成系统</div>', unsafe_allow_html=True)

with st.sidebar:
    st.header("⚙️ 系统配置")
    api_key = st.text_input("Claude API Key", type="password")
    st.markdown("---")
    amap_key = st.text_input("高德地图 API Key（可选）", type="password",
                              help="用于自动获取周边配套信息")
    st.markdown("---")
    st.markdown("**使用步骤：**\n1. 配置API Key\n2. 上传PDF\n3. 提取信息\n4. 确认字段\n5. 生成报告")
    st.markdown("---")
    if TEMPLATE_PATH.exists():
        st.success("✅ 模板已就绪")
    else:
        st.error("❌ 未找到 template_v2.docx")

# 上传区
col1, col2 = st.columns([1, 1], gap="large")

with col1:
    st.subheader("📤 上传 PDF 文件")
    uploaded_pdf = st.file_uploader("支持扫描件PDF", type=["pdf"])
    if uploaded_pdf:
        st.success(f"✅ {uploaded_pdf.name}（{uploaded_pdf.size/1024:.1f} KB）")
        if not api_key:
            st.warning("⚠️ 请在左侧填写 Claude API Key")
        else:
            if st.button("🚀 开始提取信息", type="primary", use_container_width=True):
                with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                    tmp.write(uploaded_pdf.read())
                    tmp_path = tmp.name
                try:
                    extracted = extract_info_from_pdf(tmp_path, api_key)
                    st.session_state["extracted"] = extracted
                    st.session_state["extraction_done"] = True
                    st.success("✅ 提取完成！请在右侧确认字段。")
                except Exception as e:
                    st.error(f"❌ 提取失败：{e}")
                finally:
                    os.unlink(tmp_path)

with col2:
    st.subheader("📋 提取结果预览")
    if st.session_state.get("extraction_done"):
        for k, v in st.session_state["extracted"].items():
            if v and v != "未提及":
                st.markdown(f"**{k}：** {v}")
    else:
        st.info("上传PDF提取后，结果显示在这里")


# ===================== 编辑表单 & 生成 =====================
if st.session_state.get("extraction_done"):
    st.markdown("---")
    st.subheader("✏️ 确认 & 编辑字段")
    extracted = st.session_state["extracted"]

    with st.form("report_form"):
        edited = {}

        for group_name, fields in FIELD_GROUPS.items():
            st.markdown(f"**{group_name}**")
            cols = st.columns(3)
            for i, field in enumerate(fields):
                with cols[i % 3]:
                    edited[field] = st.text_input(
                        field, value=extracted.get(field, ""), key=f"f_{field}"
                    )
            st.markdown("")

        # 图片上传
        st.markdown("---")
        st.markdown("**🖼️ 图片上传（可选）**")
        img_cols = st.columns(5)
        img_labels = {
            "IMAGE_LOCATION_MAP": "位置示意图",
            "IMAGE_PHOTO_1":      "实景照片 1",
            "IMAGE_PHOTO_2":      "实景照片 2",
            "IMAGE_PHOTO_3":      "实景照片 3",
            "IMAGE_PHOTO_4":      "实景照片 4",
        }
        uploaded_images = {}
        for i, (img_key, img_label) in enumerate(img_labels.items()):
            with img_cols[i]:
                f = st.file_uploader(img_label, type=["jpg", "jpeg", "png"], key=f"img_{img_key}")
                if f:
                    uploaded_images[img_key] = f

        st.markdown("---")
        col_a, col_b = st.columns(2)
        with col_a:
            fetch_surr = st.checkbox(
                "🗺️ 自动获取周边配套并生成区位描述",
                value=bool(amap_key),
                disabled=not bool(amap_key),
                help="需填写高德地图API Key"
            )
        with col_b:
            submitted = st.form_submit_button(
                "📝 生成Word报告", type="primary", use_container_width=True
            )

        if submitted:
            surroundings = {}
            para1, para2 = "", ""
            address = edited.get("房产地址", "")

            # 获取周边配套
            if fetch_surr and amap_key and address:
                with st.spinner("🗺️ 正在搜索周边配套..."):
                    surroundings = search_surroundings(address, amap_key)

            # 生成区位描述
            if api_key and address:
                with st.spinner("✍️ AI正在生成区位描述..."):
                    para1, para2 = generate_surrounding_description(edited, surroundings, api_key)

            # 构建填充数据（字段 + 区位描述）
            fill_data = dict(edited)
            fill_data["区位描述段1"] = para1
            fill_data["区位描述段2"] = para2

            output_path = os.path.join(tempfile.gettempdir(), f"评估报告_{int(time.time())}.docx")
            with st.spinner("📄 正在填充模板生成报告..."):
                try:
                    fill_template(fill_data, output_path)

                    # 替换图片
                    for img_key, img_file in uploaded_images.items():
                        ext = img_file.name.rsplit(".", 1)[-1].lower()
                        replace_image_in_docx(output_path, img_key, img_file.read(), ext)

                    with open(output_path, "rb") as f:
                        doc_bytes = f.read()

                    owner = edited.get("权属人", "报告") or "报告"
                    addr_short = (edited.get("房产地址") or "")[:10]
                    filename = f"{owner}_{addr_short}_估价报告.docx"

                    # 保存到session_state，在表单外显示下载按钮
                    st.session_state["report_bytes"] = doc_bytes
                    st.session_state["report_filename"] = filename
                    st.session_state["report_para1"] = para1
                    st.session_state["report_para2"] = para2
                    st.success("✅ 报告生成成功！请点击下方按钮下载。")

                except FileNotFoundError as e:
                    st.error(str(e))
                except Exception as e:
                    st.error(f"❌ 报告生成失败：{e}")
                    st.code(str(e))

# ===================== 下载按钮（表单外）=====================
if st.session_state.get("report_bytes"):
    st.download_button(
        label="⬇️ 下载Word报告",
        data=st.session_state["report_bytes"],
        file_name=st.session_state["report_filename"],
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )
    para1 = st.session_state.get("report_para1", "")
    para2 = st.session_state.get("report_para2", "")
    if para1:
        st.markdown("---")
        st.subheader("📍 区位描述预览")
        st.markdown(
            f'<div class="section-card">{para1}<br><br>{para2}</div>',
            unsafe_allow_html=True
        )
