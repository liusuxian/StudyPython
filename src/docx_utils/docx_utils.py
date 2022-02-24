import docx
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE


# 定义 创建 超链接 函数
def addHyperlink(paragraph, url, text, color, underline):
    """
    A function that places a hyperlink within a paragraph object.
    :param paragraph: The paragraph we are adding the hyperlink to.
    :param url: A string containing the required url
    :param text: The text displayed for the url
    :param color:
    :param underline:
    :return: The hyperlink object
    """
    # This gets access to the document.xml.rels file and gets a new relation id value
    part = paragraph.part
    r_id = part.relate_to(url, RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    # Create the w:hyperlink tag and add needed values
    hyperLink = docx.oxml.shared.OxmlElement('w:hyperlink')
    hyperLink.set(docx.oxml.shared.qn('r:id'), r_id)
    # Create a w:r element
    new_run = docx.oxml.shared.OxmlElement('w:r')
    # Create a new w:rPr element
    rPr = docx.oxml.shared.OxmlElement('w:rPr')
    # Add color if it is given
    if not color is None:
        c = docx.oxml.shared.OxmlElement('w:color')
        c.set(docx.oxml.shared.qn('w:val'), color)
        rPr.append(c)
    # Remove underlining if it is requested
    if not underline:
        u = docx.oxml.shared.OxmlElement('w:u')
        u.set(docx.oxml.shared.qn('w:val'), 'none')
        rPr.append(u)
    # Join all the xml elements together add add the required text to the w:r element
    new_run.append(rPr)
    new_run.text = text
    hyperLink.append(new_run)
    paragraph._p.append(hyperLink)

    return hyperLink


# 定义 创建 书签 函数
def addBookmark(paragraph, bookmark_text, bookmarkName):
    """
    A function that places a hyperlink within a paragraph object.
    :param paragraph: 创建一个段落
    :param bookmark_text: 定位到书签文本会被插入到文档中，
    :param bookmarkName: 书签名称
    """
    run = paragraph.add_run()
    tag = run._r  # for reference the following also works: tag =  document.element.xpath('//w:r')[-1]
    start = docx.oxml.shared.OxmlElement('w:bookmarkStart')
    start.set(docx.oxml.ns.qn('w:id'), '0')
    start.set(docx.oxml.ns.qn('w:name'), bookmarkName)
    tag.append(start)

    text = docx.oxml.OxmlElement('w:r')
    text.text = bookmark_text
    tag.append(text)

    end = docx.oxml.shared.OxmlElement('w:bookmarkEnd')
    end.set(docx.oxml.ns.qn('w:id'), '0')
    end.set(docx.oxml.ns.qn('w:name'), bookmarkName)
    tag.append(end)
