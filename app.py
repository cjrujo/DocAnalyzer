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
import traceback
import logging
import tempfile
import mimetypes

# -*- coding: utf-8 -*-

# 加载.env文件
load_dotenv()

# 初始化会话状态变量
for key in ['analysis_result', 'combined_content']:
    if key not in st.session_state:
        st.session_state[key] = None

def get_api_key():
    api_key = st.secrets.get("GOOGLE_API_KEY")
    if not api_key:
        api_key = st.sidebar.text_input("输入您的Google Gemini API密钥", type="password")
        if api_key:
            st.sidebar.success("API密钥已设置并保存!")
    else:
        st.sidebar.success("已从Secrets中读取API密钥")
    return api_key

st.title("文档分析器")

api_key = get_api_key()
max_tokens = st.slider("选择要处理的最大文本长度 (token)", 1000, 1000000, 128000)
temperature = st.slider("选择模型的temperature", 0.0, 1.0, 0.3, 0.1)

def try_gemini_file_api(file, selected_prompt):
    logging.info(f"开始使用Gemini API处理内容")
    logging.info(f"API密钥前缀: {api_key[:5]}...")
    logging.info(f"使用的Gemini模型: gemini-1.5-flash-latest")
    try:
        genai.configure(api_key=api_key)
        model = genai.GenerativeModel('gemini-1.5-flash-latest')
        
        # 准备内容
        if isinstance(file, BytesIO):
            file_content = file.getvalue()
        elif isinstance(file, str):
            file_content = file.encode('utf-8')
        elif hasattr(file, 'read'):
            file_content = file.read()
        else:
            raise ValueError(f"Unsupported file type: {type(file)}")

        mime_type = getattr(file, 'type', 'application/octet-stream')
        if mime_type == 'application/octet-stream' and hasattr(file, 'name'):
            guessed_type = mimetypes.guess_type(file.name)[0]
            if guessed_type:
                mime_type = guessed_type

        # 处理Excel文件
        if mime_type == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
            logging.info(f"处理Excel文件")
            text_content = extract_content_from_xlsx(BytesIO(file_content))
            prompt_parts = [selected_prompt, text_content]
        elif mime_type.startswith('image/'):
            logging.info(f"处理图像文件")
            image = Image.open(BytesIO(file_content))
            prompt_parts = [selected_prompt, image]
        elif mime_type in ['text/plain', 'text/html', 'text/csv', 'application/json']:
            logging.info(f"处理文本文件")
            text_content = file_content.decode('utf-8')
            prompt_parts = [selected_prompt, text_content]
        elif mime_type == 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
            logging.info(f"处理Word文档")
            # 使用docx库提取文本
            doc = Document(BytesIO(file_content))
            text_content = "\n".join([paragraph.text for paragraph in doc.paragraphs])
            prompt_parts = [selected_prompt, text_content]
        else:
            logging.info(f"处理其他类型文件: {mime_type}")
            # 创建临时文件
            with tempfile.NamedTemporaryFile(delete=False, suffix=f".{mime_type.split('/')[-1]}") as temp_file:
                temp_file.write(file_content)
                temp_file_path = temp_file.name

            try:
                uploaded_file = genai.upload_file(temp_file_path, mime_type=mime_type)
                prompt_parts = [selected_prompt, uploaded_file]
            finally:
                # 删除临时文件
                os.unlink(temp_file_path)

        logging.info("开始调用Gemini API")
        response = model.generate_content(
            prompt_parts,
            generation_config=genai.types.GenerationConfig(temperature=temperature)
        )
        logging.info("Gemini API调用成功")
        return response.text
    except Exception as e:
        logging.error(f"Gemini 文件 API 处理过程中发生错误：{str(e)}", exc_info=True)
        return f"Gemini 文件 API 处理过程中发生错误：{str(e)}"

def extract_content_from_file(file):
    content = try_gemini_file_api(file, "请分析这个文件的内容")
    if content is not None:
        return content
    
    # 如果 Gemini API 失败, 使用原始方法
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
    if extractor:
        return extractor(file)
    else:
        return {'type': 'error', 'data': f"不支持的文件类型: {file_extension}"}

def process_image(image, caption):
    buffered = BytesIO()
    image.save(buffered, format="PNG")
    img_str = base64.b64encode(buffered.getvalue()).decode()
    return {'type': 'image', 'data': img_str, 'caption': caption}

def extract_content_from_pdf(pdf_file):
    content = []
    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                content.append(text)
            for image in page.images:
                try:
                    img = Image.open(BytesIO(image['stream'].get_data()))
                    content.append(process_image(img, f"PDF中的图片 (页面 {page.page_number})"))
                except Exception as e:
                    print(f"无法处理PDF中的图片 (页面 {page.page_number}): {str(e)}")
                    content.append(f"[无法处理的图片] (页面 {page.page_number})")
    return content

def extract_content_from_pptx(file):
    content = []
    prs = Presentation(file)
    for slide in prs.slides:
        slide_content = []
        for shape in slide.shapes:
            if hasattr(shape, 'text'):
                slide_content.append(shape.text)
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                try:
                    image = Image.open(BytesIO(shape.image.blob))
                    content.append(process_image(image, f"PPT中的图片 (幻灯片 {slide.slide_id})"))
                except Exception as e:
                    print(f"无法处理PPT中的图片 (幻灯片 {slide.slide_id}): {str(e)}")
                    content.append(f"[无法处理的图片] (幻灯片 {slide.slide_id})")
        if slide_content:
            content.append(" ".join(slide_content))
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
    content = []
    doc = Document(file)
    for block in iter_block_items(doc):
        if isinstance(block, Paragraph):
            content.append(block.text)
        elif isinstance(block, Table):
            for row in block.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        content.append(paragraph.text)
    return content

def extract_content_from_xlsx(file):
    content = []
    workbook = openpyxl.load_workbook(file)
    for sheet in workbook.sheetnames:
        worksheet = workbook[sheet]
        for row in worksheet.iter_rows(values_only=True):
            content.append(" | ".join(str(cell) for cell in row if cell is not None))
    return content

def extract_content_from_text(file):
    content = file.read().decode('utf-8')
    if file.name.endswith('.md'):
        content = markdown.markdown(content)
    return [content]

def extract_content_from_image(file):
    image = Image.open(file)
    return [process_image(image, f"上传的图片: {file.name}")]

def get_prompt_options():
    return {
        "选项1 - 综合分析": """
        请按照以下步骤分析给定的内容,包括文本和图像(如果有):

        1. 仔细阅读/观察所有提供的内容。
        2. 识别主要主题和关键点。
        3. 分析内容的整体语气和情感。
        4. 如果有图像,详细描述并分析其内容。
        5. 考虑内容的潜在目的或意图。
        6. 注意任何独特或有趣的方面。
        7. 总结你的发现,提供一个全面的分析。

        基于以上步骤,请提供以下分析:

        1. 主要内容概述
        2. 关键点或主题列表
        3. 情感或语气分析
        4. 图像描述和分析(如果适用)
        5. 内容的潜在目的或意图
        6. 任何值得注意的独特或有趣的方面

        请首先用英语提供你的分析,然后提供相同内容的中文翻译。确保两种语言的内容一致。

        以下是需要分析的内容:

        {combined_content}
        """,
        
        "选项2 - 摘要生成": """
        请按照以下步骤为给定内容生成摘要:

        1. 仔细阅读整个文本。
        2. 识别主要论点或关键信息。
        3. 删除不必要的细节和例子。
        4. 用自己的话重新表述主要观点。
        5. 确保摘要涵盖原文的所有重要方面。
        6. 保持摘要简洁,不超过原文的25%。

        基于以上步骤,请生成一个简洁的摘要。

        请首先用英语提供摘要,然后提供相同内容的中文翻译。确保两种语言的内容一致。

        以下是需要总结的内容:

        {combined_content}
        """,
        
        "选项3 - 关键词提取": """
        请按照以下步骤从给定内容中提取关键词或短语:

        1. 仔细阅读整个文本。
        2. 识别频繁出现或对理解内容至关重要的词语。
        3. 考虑专业术语或特定领域的词汇。
        4. 注意可能概括主题的短语。
        5. 选择5-10个最能代表内容核心的词或短语。

        基于以上步骤,请提取5-10个关键词或短语。

        请首先用英语列出这些关键词/短语,然后提供相同内容的中文翻译。确保两种语言的内容一致。

        以下是需要提取关键词的内容:

        {combined_content}
        """,
        
        "选项4 - 情感分析": """
        请按照以下步骤分析给定内容的情感倾向:

        1. 仔细阅读整个文本。
        2. 识别带有情感色彩的词语和短语。
        3. 考虑整体语气和作者的态度。
        4. 评估正面和负面表达的比例。
        5. 考虑任何中性或客观的陈述。
        6. 根据以上观察,判断整体情感倾向(积极、消极或中性)。

        基于以上步骤,请分析内容的整体情感倾向,并提供支持你结论的具体例子。

        请首先用英语提供你的分析,然后提供相同内容的中文翻译。确保两种语言的内容一致。

        以下是需要进行情感分析的内容:

        {combined_content}
        """,
        
        "选项5 - 图像分析": """
        请按照以下步骤分析给定的图像:

        1. 仔细观察图像的所有元素。
        2. 描述图像中的主要对象、人物或场景。
        3. 注意颜色、光线、构图等视觉元素。
        4. 分析图像可能传达的情感或氛围。
        5. 考虑图像的潜在含义或象征意义。
        6. 如果适用,讨论图像与其上下文的关系。

        基于以上步骤,请提供一个详细的图像分析。

        请首先用英语提供你的分析,然后提供相同内容的中文翻译。确保两种语言的内容一致。

        以下是需要分析的图像内容:

        {combined_content}
        """,
        
        "选项6 - 简报制作": """
        请按照以下步骤为给定内容制作一份简报:

        1. 仔细阅读所有提供的内容。
        2. 确定简报的主题和目标受众。
        3. 提取3-5个关键信息点。
        4. 为每个关键点设计一个简洁的标题。
        5. 为每个关键点准备支持数据或例子。
        6. 考虑如何以视觉方式呈现信息(图表、图像等)。
        7. 设计一个引人注目的开场和结束语。
        8. 组织内容的逻辑流程。
        9. 考虑可能的问答环节,准备潜在问题的答案。

        基于以上步骤,请提供以下内容:

        1. 简报标题
        2. 目标受众
        3. 3-5个关键信息点(每个点包括标题和简要说明)
        4. 每个关键点的支持数据或例子
        5. 建议的视觉元素(如图表类型、图像描述等)
        6. 开场白
        7. 结束语
        8. 3-5个可能的问答问题及其答案

        请首先用英语提供你的简报内容,然后提供相同内容的中文翻译。确保两种语言的内容一致。

        以下是需要制作简报的内容:

        {combined_content}
        """
    }

def process_files(uploaded_files, selected_prompt):
    contents = []
    for file in uploaded_files:
        logging.info(f"开始处理文件: {file.name}")
        with st.spinner(f'正在处理文件 {file.name}...'):
            try:
                content = extract_content_from_file(file)
                if isinstance(content, str) and content.startswith("Gemini API调用失败"):
                    st.warning(f"文件 {file.name} 处理警告: {content}")
                    logging.warning(f"文件 {file.name} 处理警告: {content}")
                else:
                    contents.extend(content if isinstance(content, list) else [content])
                    st.success(f"文件 {file.name} 处理完成")
                    logging.info(f"文件 {file.name} 处理完成")
            except Exception as e:
                logging.error(f"处理文件 {file.name} 时发生错误", exc_info=True)
                st.error(f"处理文件 {file.name} 时发生错误: {str(e)}")
    
    logging.info("开始内容分析")
    with st.spinner('正在分析内容...'):
        combined_content = "\n\n".join(str(c) for c in contents if not isinstance(c, dict))
        analysis_result = try_gemini_file_api(
            combined_content,  # 直接传递字符串内容
            selected_prompt.format(combined_content=combined_content)
        )
    
    return contents, analysis_result

prompt_options = get_prompt_options()
selected_option = st.selectbox("选择分析方式", list(prompt_options.keys()) + ["自定义提示"])
custom_prompt = st.text_area("输入自定义提示", height=200) if selected_option == "自定义提示" else prompt_options[selected_option]

uploaded_files = st.file_uploader("上传文件", type=["pdf", "pptx", "docx", "xlsx", "txt", "md", "jpg", "jpeg", "png", "gif", "bmp"], accept_multiple_files=True)

if uploaded_files:
    st.write(f"已上传 {len(uploaded_files)} 个文件")
    show_combined_content = st.checkbox("显示合并后的内容")
    
    if st.button("分析所有文件"):
        logging.info("开始分析文件")
        if not api_key:
            logging.warning("未提供API密钥")
            st.error("请先输入有效的Google Gemini API密钥")
        else:
            try:
                selected_prompt = custom_prompt if selected_option == "自定义提示" else prompt_options[selected_option]
                logging.info(f"选择的分析选项: {selected_option}")
                
                contents, analysis_result = process_files(uploaded_files, selected_prompt)
                
                if show_combined_content:
                    st.write("合并后的内容:")
                    for content in contents:
                        st.write(content)
                
                if analysis_result:
                    logging.info("分析完成, 显示结果")
                    st.session_state.analysis_result = analysis_result
                    st.session_state.combined_content = "\n\n".join(str(c) for c in contents)
                    
                    st.write("分析结果:")
                    st.write(analysis_result)
                    
                    st.download_button(
                        label="下载分析结果",
                        data=analysis_result.encode('utf-8'),
                        file_name="analysis_result.txt",
                        mime="text/plain"
                    )
                else:
                    logging.warning("分析未产生结果")
                    st.error("分析未产生结果。请检查文件内容和API密钥。")
            except Exception as e:
                logging.error("主程序中发生错误", exc_info=True)
                st.error(f"分析过程中发生错误: {str(e)}")

        logging.info("分析过程结束")

st.write("使用说明:")
st.write("1. 输入您的Google Gemini API密钥")
st.write("2. 上传文件 (支持PDF、PPTX、DOCX、XLSX、TXT、MD、JPG、JPEG、PNG、GIF和BMP格式)")
st.write("3. 选择分析方式")
st.write("4. 点击'分析所有文件'按钮")
st.write("5. 查看分析结果")

st.markdown("---")
st.write("使用gemini-1.5-flash-latest API进行文档内容分析")

logging.basicConfig(filename='app.log', level=logging.DEBUG, 
                    format='%(asctime)s %(levelname)s:%(message)s')

# 添加控制台日志处理器
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.DEBUG)
logging.getLogger('').addHandler(console_handler)

def log_exception(exc_type, exc_value, exc_traceback):
    logging.error("Uncaught exception", exc_info=(exc_type, exc_value, exc_traceback))

import sys
sys.excepthook = log_exception
