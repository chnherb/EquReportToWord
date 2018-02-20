/************************************************************************/
/* Word.cs-Word操作类
 * 
 * author:杨习辉
 * email: ahhuiyang@gmail.com
 * date:  2011-4-6
/************************************************************************/

using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Reflection;

using Word = Microsoft.Office.Interop.Word;

namespace CSAutoWord
{
    public class CSWord
    {
        private object missing = Type.Missing;
        private Word.Application m_WordApp = null;
        private Word.Documents m_WordDocs = null;
        private Word.Document m_WordDoc = null;
        private Word.Paragraph m_ContentsPara = null;

        //标题样式
        public static string WORD_HEADING1 = "标题 1";
        public static string WORD_HEADING2 = "标题 2";
        public static string WORD_HEADING3 = "标题 3";

        public CSWord()
        {
            //创建Word
            CreateWordApp(false);
            CreateWordDoc();
        }
       
        #region 插入文字

        public bool InsertText(string text, float fontsize,
            int fontbold, Word.WdParagraphAlignment align)
        {
            bool bRet = false;

            try
            {
                Word.Paragraph paragrah = m_WordDoc.Paragraphs.Add(ref missing);
                paragrah.Range.Text = text;
                paragrah.Range.Font.Size = fontsize;
                paragrah.Range.Font.Bold = fontbold;
                paragrah.Alignment = align;
                paragrah.Range.InsertParagraphAfter();

                bRet = true;
            }
            catch (System.Exception)
            {
                bRet = false;
            }

            return bRet;
        }

        //设置固定行间距lineSpacing磅
        public bool InsertText(string text, float fontsize,
            int fontbold, Word.WdParagraphAlignment align, int lineSpacing)
        {
            bool bRet = false;

            try
            {
                Word.Paragraph paragrah = m_WordDoc.Paragraphs.Add(ref missing);
                paragrah.Range.Text = text;
                paragrah.Range.Font.Size = fontsize;
                paragrah.Range.Font.Bold = fontbold;
                paragrah.Alignment = align;

                paragrah.LineSpacing = lineSpacing; //设置文档的行间距
                paragrah.LineSpacingRule = Microsoft.Office.Interop.Word.WdLineSpacing.wdLineSpaceExactly;

                paragrah.Range.InsertParagraphAfter();

                bRet = true;
            }
            catch (System.Exception)
            {
                bRet = false;
            }

            return bRet;
        }
        #endregion

        #region 插入文字，并插入一个分页符
        //插入文字
        public bool InsertTextNewPage(string text, float fontsize,
            int fontbold, Word.WdParagraphAlignment align)
        {
            bool bRet = false;

            try
            {
                Word.Paragraph paragrah = m_WordDoc.Paragraphs.Add(ref missing);
                paragrah.Range.Text = text;
                paragrah.Range.Font.Size = fontsize;
                paragrah.Range.Font.Bold = fontbold;
                paragrah.Alignment = align;
                paragrah.Range.InsertParagraphAfter();

                object pBreak = (int)Word.WdBreakType.wdSectionBreakNextPage;
                paragrah.Range.InsertBreak(ref pBreak);

                bRet = true;
            }
            catch (System.Exception)
            {
                bRet = false;
            }

            return bRet;
        }
        #endregion

        #region 插入一级标题，二级标题...
        public bool InsertText(string text,string heading)
        {
            bool bRet = false;

            try
            {
                Word.Paragraph paragrah = m_WordDoc.Paragraphs.Add(ref missing);
                paragrah.Range.Text = text;
                object obj = heading;
                paragrah.set_Style(ref obj);
                paragrah.Range.InsertParagraphAfter();

                bRet = true;
            }
            catch (System.Exception)
            {
                bRet = false;
            }

            return bRet;
        }

        public bool InsertText(string text, string heading, Word.WdParagraphAlignment align)
        {
            bool bRet = false;

            try
            {
                Word.Paragraph paragrah = m_WordDoc.Paragraphs.Add(ref missing);
                paragrah.Range.Text = text;
                object obj = heading;
                paragrah.set_Style(ref obj);
                paragrah.Alignment = align;
                paragrah.Range.InsertParagraphAfter();

                bRet = true;
            }
            catch (System.Exception)
            {
                bRet = false;
            }

            return bRet;
        }
        #endregion

        #region 插入文字，并且在文字后预留目录位置
        public bool InsertText(string text, float fontsize,
            int fontbold, Word.WdParagraphAlignment align,bool bCreateContentsPara)
        {
            bool bRet = false;

            bRet = InsertText(text, fontsize, fontbold, align);

            if (bCreateContentsPara)
            {
                m_ContentsPara = m_WordDoc.Paragraphs.Add(ref missing);
            }
            
            return bRet;
        }
        #endregion

        #region 插入一行
        public bool InsertLine()
        {
            return InsertText(string.Empty, 12, 0, Word.WdParagraphAlignment.wdAlignParagraphLeft);
        }
        #endregion

        #region 插入新页
        public bool InsertNewPage()
        {
            return InsertTextNewPage(string.Empty, 12, 0, Word.WdParagraphAlignment.wdAlignParagraphLeft);
        }
        #endregion

        #region 插入一幅图片
        public bool InsertPicture(string img, int width, int height,
            Word.WdParagraphAlignment align)
        {
            bool bRet = false;

            try
            {
                Word.Paragraph paragraph = m_WordDoc.Paragraphs.Add(ref missing);
                paragraph.Alignment = align;
                object range = paragraph.Range;
                Word.InlineShape il = paragraph.Range.InlineShapes.AddPicture(img, ref missing, ref missing, ref range);
                il.Width = width;
                il.Height = height;
                paragraph.Range.InsertParagraphAfter();

                bRet = true;
            }
            catch (System.Exception)
            {
                bRet = false;
            }

            return bRet;
        }
        #endregion

        #region 插入表格
        public Word.Table InsertTable(Word.WdParagraphAlignment align,int rows,int cols,bool bInsertNewPage)
        {
            Word.Table table = null;

            Word.Paragraph paragraph = m_WordDoc.Paragraphs.Add(ref missing);
            paragraph.Alignment = align;

            object start = 0;
            object end = 0;
            table = m_WordDoc.Tables.Add(paragraph.Range, rows, cols, ref missing, ref missing);

            if (bInsertNewPage)
            {
                object pBreak = (int)Word.WdBreakType.wdSectionBreakNextPage;
                paragraph.Range.InsertBreak(ref pBreak);
            }

            return table;
        }
        #endregion

        #region 创建目录
        public bool CreateContents()
        {
            bool bRet = false;

            if (m_ContentsPara != null)
            {
                bRet = true;
            }

            Object oTrue = true;
            Object oFalse = false;
            object x = 0;

            Word.Range myRange = m_ContentsPara.Range;

            Object oUpperHeadingLevel = "1";
            Object oLowerHeadingLevel = "2";
            Object oTOCTableID = "TableOfContents";

            m_WordDoc.TablesOfContents.Add(myRange, ref oTrue, ref oUpperHeadingLevel,
                ref oLowerHeadingLevel, ref missing, ref oTOCTableID, ref oTrue,
                ref oTrue, ref missing, ref oTrue, ref oTrue, ref oTrue);

            return bRet;
        }
       
        public bool CreateContents(bool bInsertNewPage)
        {
            bool bRet = false;

            if (m_ContentsPara != null)
            {
                bRet = true;
            }

            Object oTrue = true;
            Object oFalse = false;
            object x = 0;

            Word.Range myRange = m_ContentsPara.Range;

            Object oUpperHeadingLevel = "1";
            Object oLowerHeadingLevel = "2";
            Object oTOCTableID = "TableOfContents";

            m_WordDoc.TablesOfContents.Add(myRange, ref oTrue, ref oUpperHeadingLevel,
                ref oLowerHeadingLevel, ref missing, ref oTOCTableID, ref oTrue,
                ref oTrue, ref missing, ref oTrue, ref oTrue, ref oTrue);

            if (bInsertNewPage)
            {
                object pBreak = (int)Word.WdBreakType.wdSectionBreakNextPage;
                m_ContentsPara.Range.InsertBreak(ref pBreak);
            }

            return bRet;
        }
        #endregion

        #region 添加页眉
        public void AddSimpleHeader(string HeaderText)
        {
            //添加页眉     
            m_WordApp.ActiveWindow.View.Type = Word.WdViewType.wdOutlineView;
            m_WordApp.ActiveWindow.View.SeekView = Word.WdSeekView.wdSeekPrimaryHeader;
            m_WordApp.ActiveWindow.ActivePane.Selection.InsertAfter(HeaderText);
            m_WordApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;//设置左对齐     
            m_WordApp.ActiveWindow.View.SeekView = Word.WdSeekView.wdSeekMainDocument;
        }

        public void AddSimpleHeader(string HeaderText, Word.WdParagraphAlignment wdAlign)
        {
            //添加页眉     
            m_WordApp.ActiveWindow.View.Type = Word.WdViewType.wdOutlineView;
            m_WordApp.ActiveWindow.View.SeekView = Word.WdSeekView.wdSeekPrimaryHeader;
            m_WordApp.ActiveWindow.ActivePane.Selection.InsertAfter(HeaderText);   
            m_WordApp.Selection.ParagraphFormat.Alignment = wdAlign;//设置左对齐     
            m_WordApp.ActiveWindow.View.SeekView = Word.WdSeekView.wdSeekMainDocument;
        }

        public void AddSimpleHeader(string HeaderText, Word.WdParagraphAlignment wdAlign, Word.WdColor fontcolor, float fontsize)
        {
            //添加页眉     
            m_WordApp.ActiveWindow.View.Type = Word.WdViewType.wdOutlineView;
            m_WordApp.ActiveWindow.View.SeekView = Word.WdSeekView.wdSeekPrimaryHeader;
            m_WordApp.ActiveWindow.ActivePane.Selection.InsertAfter(HeaderText);
            m_WordApp.Selection.Font.Color = fontcolor;//设置字体颜色     
            m_WordApp.Selection.Font.Size = fontsize;//设置字体大小     
            m_WordApp.Selection.ParagraphFormat.Alignment = wdAlign;//设置对齐方式     
            m_WordApp.ActiveWindow.View.SeekView = Word.WdSeekView.wdSeekMainDocument;
        }
        #endregion

        #region 插入图片
        public void AddPicHeader(string img, Word.WdParagraphAlignment align)
        {
            m_WordDoc.ActiveWindow.ActivePane.View.SeekView = Word.WdSeekView.wdSeekPrimaryHeader;
            Word.InlineShape il = m_WordDoc.ActiveWindow.ActivePane.Selection.InlineShapes.AddPicture(
                img, ref missing, ref missing, ref missing);

            il.Width = 20;
            il.Height = 20;

            il.HorizontalLineFormat.Alignment = Word.WdHorizontalLineAlignment.wdHorizontalLineAlignRight;
        }
        #endregion

        #region 创建WordApp
        public Word.Application CreateWordApp(bool visible)
        {
            try
            {
                m_WordApp = new Word.Application();
                m_WordApp.Visible = visible;
            }
            catch (System.Exception)
            {
                m_WordApp = null;
            }

            return m_WordApp;
        }

        public Word.Application CreateWordApp()
        {
            try
            {
                m_WordApp = new Word.Application();
                m_WordApp.Visible = false;
            }
            catch (System.Exception)
            {
                m_WordApp = null;
            }

            return m_WordApp;
        }
        #endregion

        #region 创建WordDoc
        public Word.Document CreateWordDoc()
        {
            try
            {
                //新建一个文档
                m_WordDocs = m_WordApp.Documents;
                m_WordDoc = m_WordDocs.Add(ref missing, ref missing, ref missing, ref missing);
            }
            catch (System.Exception)
            {
                m_WordDoc = null;
            }

            return m_WordDoc;
        }
        #endregion

        #region 保存文档
        public bool SaveWordDocument(string path,string filename)
        {
            bool bRet = false;

            try
            {
                path = path.Trim();
                if (path == string.Empty)
                {
                    path = Assembly.GetExecutingAssembly().Location;
                }
                object FileName = path + "\\" + filename + ".doc";
                object FileFormat = Word.WdSaveFormat.wdFormatDocument;
                m_WordDoc.SaveAs(ref FileName, ref FileFormat, ref missing,
                        ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing,
                        ref missing);
                ((Word._Document)m_WordDoc).Close(ref missing, ref missing, ref missing);
                ((Word._Application)m_WordApp).Quit(ref missing, ref missing, ref missing);

                bRet = true;
            }
            catch (System.Exception)
            {
                bRet = false;
            }

            return bRet;
        }
        #endregion

        #region 打开文档
        public bool DisplayWordFile(object filePath)//形如"D:/文档",省略".doc"
        {
            filePath += ".doc";
            try
            {
                Word.Application G_wa = new Microsoft.Office.Interop.Word.Application();
                G_wa.Visible = true;

                Word.Document P_Document = G_wa.Documents.Open(ref filePath, ref missing,
                                                              ref missing, ref missing, ref missing, ref missing,
                                                              ref missing, ref missing, ref missing, ref missing,
                                                              ref missing, ref missing, ref missing, ref missing,
                                                              ref missing, ref missing);
            }
            catch (Exception e)
            {
                return false;
            }
            return true;
        }
        #endregion

    }
}