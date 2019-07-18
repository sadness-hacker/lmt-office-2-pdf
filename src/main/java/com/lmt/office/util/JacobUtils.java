package com.lmt.office.util;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.ComThread;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import lombok.extern.slf4j.Slf4j;

import java.io.File;

/**
 * @description word/excel/ppt转pdf工具类,基于Jacob实现,需要运行在有MS office套件的Windows电脑上
 * @author bazhandao
 * @date 2019-07-17
 */
@Slf4j
public class JacobUtils {

    // Word转 PDF 格式
    private static final int WORD_FORMAT_PDF = 17;
    // Excel转 PDF 格式
    private static final int EXCEL_FORMAT_PDF = 0;
    // PPT转 PDF 格式
    private static final int PPT_FORMAT_PDF = 32;

    /**
     * word/excel/ppt转pdf工具类
     * @author bazhandao
     * @date 2019-07-17
     * @param filePath     源文件路径
     * @param pdfFilePath  pdf文件路径
     */
    public static void file2pdf(String filePath, String pdfFilePath) {
        String suffix = getFileSuffix(filePath);
        if (".doc".equals(suffix) || ".docx".equals(suffix)) {
            word2pdf(filePath, pdfFilePath);
        } else if (".ppt".equals(suffix) || ".pptx".equals(suffix)) {
            ppt2pdf(filePath, pdfFilePath);
        } else if (".xls".equals(suffix) || ".xlsx".equals(suffix)) {
            excel2pdf(filePath, pdfFilePath);
        } else {
            throw new RuntimeException("Unsupport FileType " + filePath);
        }
    }

    /**
     * 获取文件扩展名
     * @author bazhandao
     * @date 2019-07-17
     * @param filePath
     * @return
     */
    private static String getFileSuffix(String filePath) {
        int dot = filePath.lastIndexOf('.');
        return filePath.substring(dot).toLowerCase();
    }

    /**
     * word转pdf文件实现
     * @author bazhandao
     * @date 2019-07-17
     * @param fileName
     * @param pdfFileName
     * @return
     */
    private static String word2pdf(String fileName, String pdfFileName) {
        log.info("启动word转换pdf...");
        long start = System.currentTimeMillis();
        ActiveXComponent app = null;
        Dispatch doc = null;
        try {
            app = new ActiveXComponent("Word.Application");
            app.setProperty("Visible", new Variant(false));
            // 打开word文件
            Dispatch docs = app.getProperty("Documents").toDispatch();
            doc = Dispatch.invoke(docs,"Open",Dispatch.Method,new Object[] { fileName, new Variant(false),new Variant(true) }, new int[1]).toDispatch();
            log.info("打开文档...{}",fileName);
            log.info("转换文档到PDF...{}", pdfFileName);
            File tofile = new File(pdfFileName);
            if (tofile.exists()) {
                tofile.delete();
            }
            // 作为html格式保存到临时文件：：new Variant(8)其中8表示word转html;7表示word转txt;44表示Excel转html;17表示word转成pdf。
            Dispatch.invoke(doc, "SaveAs", Dispatch.Method, new Object[] { pdfFileName, new Variant(WORD_FORMAT_PDF) }, new int[1]);
            long end = System.currentTimeMillis();
            log.info("转换完成,用时：" + (end - start) + "ms.");
        } catch (Exception e) {
            log.error("word转pdf出错!!!", e);
        }catch(Throwable t){
            log.error("word转pdf出错!!!", t);
        } finally {
            // 关闭word
            Dispatch.call(doc,"Close",false);
            log.info("word转pdf,关闭文档");
            if (app != null) {
                app.invoke("Quit", new Variant[]{});
            }
            //如果没有这句winword.exe进程将不会关
            ComThread.Release();
        }
        return pdfFileName;
    }

    /**
     * excel转pdf文件
     * @date 2019-07-17
     * @author bazhandao
     * @param fileName
     * @param pdfFileName
     * @return
     */
    private static String excel2pdf(String fileName, String pdfFileName) {
        log.info("启动excel转word,fileName={},pdfFileName={}", fileName, pdfFileName);
        ActiveXComponent ax = null;
        Dispatch excel = null;
        try {
            ComThread.InitSTA();
            ax = new ActiveXComponent("Excel.Application");
            ax.setProperty("Visible", new Variant(false));
            ax.setProperty("AutomationSecurity", new Variant(3)); // 禁用宏
            Dispatch excels = ax.getProperty("Workbooks").toDispatch();
            Object[] obj = new Object[]{ fileName, new Variant(false), new Variant(false) };
            excel = Dispatch.invoke(excels, "Open", Dispatch.Method, obj, new int[9]).toDispatch();
            File tofile = new File(pdfFileName);
            // System.err.println(getDocPageSize(new File(sfileName)));
            if (tofile.exists()) {
                tofile.delete();
            }
            // 转换格式
            Object[] obj2 = new Object[]{
                    new Variant(EXCEL_FORMAT_PDF), // PDF格式=0
                    pdfFileName,
                    new Variant(0)  //0=标准 (生成的PDF图片不会变模糊) ; 1=最小文件
            };
            Dispatch.invoke(excel, "ExportAsFixedFormat", Dispatch.Method, obj2, new int[1]);
            log.info("excel转pdf转换完毕！");
        } catch (Exception es) {
            log.error("excel转pdf出错!!!", es);
        } finally {
            if (excel != null) {
                Dispatch.call(excel, "Close", new Variant(false));
            }
            if (ax != null) {
                ax.invoke("Quit", new Variant[] {});
            }
            ComThread.Release();
        }
        return pdfFileName;
    }

    /**
     * ppt转pdf
     * @author bazhandao
     * @date 2019-07-17
     * @param inputFile
     * @param pdfFile
     */
    private static void ppt2pdf(String inputFile, String pdfFile) {
        log.info("启动PPT转pdf...");
        long start = System.currentTimeMillis();
        ActiveXComponent app = null;
        Dispatch ppt = null;
        try {
            // 创建一个ppt对象
            app = new ActiveXComponent("PowerPoint.Application");
            // 不可见打开（PPT转换不运行隐藏，所以这里要注释掉）
            // app.setProperty("Visible", new Variant(false));
            // 获取文挡属性
            Dispatch ppts = app.getProperty("Presentations").toDispatch();
            // 调用Documents对象中Open方法打开文档，并返回打开的文档对象Document
            ppt = Dispatch.call(ppts, "Open", inputFile, true, true, false).toDispatch();
            log.info("打开文档...{}", inputFile);
            log.info("转换文档到PDF...{}", pdfFile);
            File tofile = new File(pdfFile);
            if(tofile.exists()) {
                tofile.delete();
            }
            Dispatch.call(ppt, "SaveAs", pdfFile, PPT_FORMAT_PDF);
            long end = System.currentTimeMillis();
            log.info("转换完成...用时：" + (end - start) + "ms.");
        } catch (Exception e) {
            log.error("ppt转pdf转换出错!!!", e);
        } finally {
            Dispatch.call(ppt, "Close");
            log.info("关闭文档");
            if (app != null) {
                app.invoke("Quit", new Variant[]{});
            }
            //如果没有这句话,winword.exe进程将不会关闭
            ComThread.Release();
        }
    }

}
