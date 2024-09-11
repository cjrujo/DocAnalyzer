import streamlit as st
import google.generativeai as genai
import os
import pdfplumber
import base64
from io import BytesIO
from PIL import Image
from dotenv import load_dotenv, set_key
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from docx import Document
from docx.document import Document as _Document
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
import openpyxl
import markdown

# 加载.env文件
load_dotenv()

# 初始化会话状态变量
for key in ['analysis_result', 'combined_content']:
    if key not in st.session_state:
        st.session_state[key] = None

def get_api_key():
    api_key = os.getenv("GOOGLE_API_KEY")
    if not api_key:
        api_key = st.sidebar.text_input("输入您的Google Gemini API密钥", type="password")
        if api_key:
            set_key('.env', "GOOGLE_API_KEY", api_key)
            os.environ["GOOGLE_API_KEY"] = api_key
            st.sidebar.success("API密钥已设置并保存!")
    else:
        st.sidebar.success("已从.env文件中读取API密钥")
    return api_key

st.title("文档分析器")

api_key = get_api_key()
max_tokens = st.slider("选择要处理的最大文本长度 (token)", 1000, 1000000, 128000)
temperature = st.slider("选择模型的temperature", 0.0, 1.0, 0.3, 0.1)

def process_image(image, caption):
    buffered = BytesIO()
    image.save(buffered, format="PNG")
    img_str = base64.b64encode(buffered.getvalue()).decode()
    return {'type': 'image', 'data': img_str, 'caption': caption}

def extract_content_from_pdf(pdf_file):
    with pdfplumber.open(pdf_file) as pdf:
        content = []
        for page_num, page in enumerate(pdf.pages, 1):
            text = page.extract_text()
            if text:
                content.append(f"第 {page_num} 页文本:\n{text}\n")
            
            for i, table in enumerate(page.extract_tables(), 1):
                content.append(f"第 {page_num} 页表格 {i}:")
                content.extend(" | ".join(str(cell) for cell in row) for row in table)
                content.append("")
            
            for i, img in enumerate(page.images, 1):
                try:
                    image = Image.open(BytesIO(img['stream'].get_data()))
                    content.append(process_image(image, f"第 {page_num} 页图像 {i}"))
                except Exception as e:
                    content.append(f"第 {page_num} 页图像 {i}: [处理图像时出错: {str(e)}]")
        return content

def extract_content_from_pptx(file):
    content = []
    prs = Presentation(file)
    for slide_num, slide in enumerate(prs.slides, 1):
        slide_content = []
        for shape_num, shape in enumerate(slide.shapes, 1):
            if shape.has_text_frame:
                slide_content.extend(paragraph.text for paragraph in shape.text_frame.paragraphs)
            elif shape.shape_type == MSO_SHAPE_TYPE.TABLE:
                table_content = [f"幻灯片 {slide_num} 表格 {shape_num}:"]
                table_content.extend(" | ".join(cell.text for cell in row.cells) for row in shape.table.rows)
                slide_content.extend(table_content)
            elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                image = Image.open(BytesIO(shape.image.blob))
                content.append(process_image(image, f"幻灯片 {slide_num} 图像 {shape_num}"))
        content.append(f"幻灯片 {slide_num} 内容:\n" + "\n".join(slide_content))
    return content

def iter_block_items(parent):
    if isinstance(parent, _Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("不支持的父元素类型")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)

def extract_content_from_docx(file):
    doc = Document(file)
    content = []
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            content.append(block.text)
        elif isinstance(block, Table):
            table_content = []
            for row in block.rows:
                table_content.append(" | ".join(cell.text for cell in row.cells))
            content.append("\n".join(table_content))

    for rel in doc.part.rels.values():
        if "image" in rel.reltype:
            image = Image.open(BytesIO(rel.target_part.blob))
            content.append(process_image(image, "文档图像"))
    return content

def extract_content_from_xlsx(file):
    wb = openpyxl.load_workbook(file)
    content = []
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        content.append(f"工作表: {sheet}")
        content.extend(" | ".join(str(cell.value) for cell in row) for row in ws.iter_rows())
        for image in ws._images:
            img = Image.open(BytesIO(image.ref))
            content.append(process_image(img, f"工作表 {sheet} 图像"))
    return content

def extract_content_from_text(file):
    content = file.getvalue().decode('utf-8')
    return {'type': 'text', 'data': markdown.markdown(content) if file.name.endswith('.md') else content}

def extract_content_from_image(file):
    try:
        return process_image(Image.open(file), "上传的图像")
    except Exception as e:
        return {'type': 'error', 'data': f"处理图像时出错: {str(e)}"}

def extract_content_from_file(file):
    extractors = {
        '.pdf': extract_content_from_pdf,
        '.pptx': extract_content_from_pptx,
        '.docx': extract_content_from_docx,
        '.xlsx': extract_content_from_xlsx,
        '.txt': extract_content_from_text,
        '.md': extract_content_from_text,
        '.jpg': extract_content_from_image,
        '.jpeg': extract_content_from_image,
        '.png': extract_content_from_image,
        '.gif': extract_content_from_image,
        '.bmp': extract_content_from_image
    }
    file_extension = os.path.splitext(file.name)[1].lower()
    extractor = extractors.get(file_extension)
    return extractor(file) if extractor else {'type': 'error', 'data': f"不支持的文件类型: {file_extension}"}

def get_prompt_options():
    return {
        "选项1 - 综合分析": """
        分析以下多个文件的内容,包括文本、表格和图像描述,并以系统化的结构总结所有文件的关键要点:

        {combined_content}

        请提供以下格式的总结:
        1. 所有文档的主要主题
        2. 关键要点 (列出5-7个要点,包括文本、表格和图像中的信息)
        3. 主要结论
        4. 建议的后续行动 (如果适用)
        5. 文档之间的关联性或差异 (如果适用)
        6. 表格和图像的重要信息摘要
        """,
        "选项2 - 创建速查表": """
        基于以下多个文件的内容,创建一个简洁的速查表:

        {combined_content}

        速查表应包含:
        1. 主要概念和定义
        2. 关键流程或步骤
        3. 重要数据或统计信息
        4. 最佳实践或建议
        5. 常见问题及解决方案
        请以易于理解和快速参考的格式呈现信息。
        """,
        "选项3 - PDCA分析": """
        使用PDCA (计划-执行-检查-行动) 方法分析以下多个文件中的问题,并提供行动计划:

        {combined_content}

        请提供以下内容:
        1. 问题识别: 列出文档中提到的主要问题或挑战
        2. 严重性分析: 对每个问题进行严重性评估 (高/中/低)
        3. 根本原因分析: 尝试确定每个问题的潜在根本原因
        4. 行动计划: 
           - 针对高严重性问题的即时行动建议
           - 针对中等严重性问题的中期行动建议
           - 针对低严重性问题的长期改进建议
        5. 监控和评估建议: 如何跟踪和衡量改进措施的效果
        """,
        "选项4 - PowerPoint演示计划": """
        基于以下文档内容,创建一个详细的PowerPoint演示文稿计划:

        {combined_content}

        请提供以下格式的PowerPoint演示计划:

        1. 标题幻灯片:
           - 主标题: [插入主标题]
           - 副标题: [插入副标题]

        2. 目录幻灯片:
           - [列出演示文稿的主要部分]

        3. 内容幻灯片 (为每个主要部分创建2-3张幻灯片):
           a. [第一部分标题]
              - 要点1: [简洁的描述]
              - 要点2: [简洁的描述]
              - 要点3: [简洁的描述]
              [建议的视觉元素: 图表、图像或图标]

           b. [第二部分标题]
              - 要点1: [简洁的描述]
              - 要点2: [简洁的描述]
              - 要点3: [简洁的描述]
              [建议的视觉元素: 图表、图像或图标]

           c. [继续添加更多部分...]

        4. 关键数据幻灯片:
           - [列出需要突出显示的重要数据或统计信息]
           [建议的视觉元素: 图表、表格或信息图]

        5. 总结幻灯片:
           - 主要结论1: [简述]
           - 主要结论2: [简述]
           - 主要结论3: [简述]

        6. 行动步骤幻灯片:
           - 建议的后续行动1: [简述]
           - 建议的后续行动2: [简述]
           - 建议的后续行动3: [简述]

        7. 结束幻灯片:
           - [感谢语]
           - [联系信息或下一步指示]

        注意:
        - 每张幻灯片应包含简洁、有影响力的内容
        - 使用简单的语言和短句
        - 建议适当的视觉元素以增强演示效果
        - 确保整个演示文稿保持一致的主题和风格
        """,
        "选项5 - 图像分析": """
        分析以下文件内容,包括文本、表格、图像和其他数据:

        {combined_content}

        对于图像内容:
        1. 描述你看到的内容
        2. 识别主要对象、场景或主题
        3. 分析图像的整体氛围或情感
        4. 提出与图像相关的任何见解或观察

        然后,提供以下格式的总结:
        1. 文本内容的主要主题和关键点
        2. 图像内容的描述和分析
        3. 文本和图像之间的任何关联或互补信息
        4. 整体结论和见解
        5. 基于所有内容的建议或后续行动
        """
    }

def analyze_content(contents, api_key, prompt_template):
    if not api_key:
        st.error("请先输入有效的Google Gemini API密钥")
        return None
    
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-1.5-pro')
    
    parts = [prompt_template]
    for content in contents:
        if isinstance(content, dict) and content['type'] == 'image':
            image = Image.open(BytesIO(base64.b64decode(content['data'])))
            parts.extend([image, content['caption']])
        else:
            parts.append(str(content))
    
    try:
        response = model.generate_content(parts, generation_config=genai.types.GenerationConfig(
            temperature=temperature
        ))
        return response.text
    except Exception as e:
        st.error(f"生成内容时发生错误: {str(e)}")
        return None

prompt_options = get_prompt_options()
selected_option = st.selectbox("选择分析方式", list(prompt_options.keys()) + ["自定义提示"])
custom_prompt = st.text_area("输入自定义提示", height=200) if selected_option == "自定义提示" else prompt_options[selected_option]

uploaded_files = st.file_uploader("上传文件", type=["pdf", "pptx", "docx", "xlsx", "txt", "md", "jpg", "jpeg", "png", "gif", "bmp"], accept_multiple_files=True)

if uploaded_files:
    st.write(f"已上传 {len(uploaded_files)} 个文件")
    show_combined_content = st.checkbox("显示合并后的内容")
    
    if st.button("分析所有文件"):
        if not api_key:
            st.error("请先输入有效的Google Gemini API密钥")
        else:
            contents = []
            for file in uploaded_files:
                content = extract_content_from_file(file)
                contents.extend(content if isinstance(content, list) else [content])
            
            if show_combined_content:
                st.write("合并后的内容:")
                for content in contents:
                    if isinstance(content, dict) and content['type'] == 'image':
                        st.image(Image.open(BytesIO(base64.b64decode(content['data']))), caption=content['caption'])
                    else:
                        st.write(content)
            
            analysis_result = analyze_content(contents, api_key, custom_prompt)
            if analysis_result:
                st.session_state.analysis_result = analysis_result
                st.session_state.combined_content = "内容已处理,包括文本和图像"
                
                st.write("分析结果:")
                st.write(analysis_result)
                
                st.download_button(
                    label="下载分析结果",
                    data=analysis_result.encode('utf-8'),
                    file_name="analysis_result.txt",
                    mime="text/plain"
                )

st.write("使用说明:")
st.write("1. 输入您的Google Gemini API密钥")
st.write("2. 上传文件 (支持PDF、PPTX、DOCX、XLSX、TXT、MD、JPG、JPEG、PNG、GIF和BMP格式)")
st.write("3. 选择分析方式")
st.write("4. 点击'分析所有文件'按钮")
st.write("5. 查看分析结果")

st.markdown("---")
st.write("使用Gemini-1.5-Flash-Latest API进行文档内容分析")