import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;

import java.io.File;

/**
 * Description:
 *
 * @author zwl
 * @version 1.0
 * @date 2021/2/16 19:03
 */
public class Words2PdfUtil {

    /**
     * 其中第44行中的 invoke（）函数中的Variant(n)参数指定另存为的文件类型（n的取值范围是0-25），他们分别是：
     * Variant(0):doc
     * Variant(1):dot
     * Variant(2-5)，Variant(7):txt
     * Variant(6):rft
     * Variant(8)，Variant(10):htm
     * Variant(9):mht
     * Variant(11)，Variant(19-22):xml
     * Variant(12):docx
     * Variant(13):docm
     * Variant(14):dotx
     * Variant(15):dotm
     * Variant(16)、Variant(24):docx
     * Variant(17):pdf
     * Variant(18):xps
     * Variant(23):odt
     * Variant(25):与Office2003与2007的转换程序相关，执行本程序后弹出一个警告框说是需要更高版本的 Microsoft Works Converter
     * 由于我计算机上没有安装这个转换器，所以不清楚此参数代表什么格式
     */
    private static final int WD_FORMAT_PDF = 17;
    // private static final int XL_FORMAT_PDF = 0;
    // private static final int PPT_FORMAT_PDF = 32;

    public static void main(String[] args) {
        String wordsFilePath = args[0];
        String pdfFilePath = args[1];
        if (wordsFilePath == null || !wordsFilePath.endsWith(".docx") || !wordsFilePath.endsWith(".DOCX")) {
            System.out.println("必须符合 java -jar xxx.jar <xxx.docx> <xxx.pdf> 的规范");
            System.exit(0);
        }

        if (pdfFilePath == null || !pdfFilePath.endsWith(".pdf") || !pdfFilePath.endsWith(".PDF")) {
            System.out.println("必须符合 java -jar xxx.jar <xxx.docx> <xxx.pdf> 的规范");
            System.exit(0);
        }

        System.out.println("开始转换...");

        long start = System.currentTimeMillis();
        boolean isSuccess = wordToPdf(wordsFilePath, pdfFilePath);
        if (isSuccess) {
            System.out.println("耗时：" + (System.currentTimeMillis() - start) + "ms");
            System.out.println("转换成功");
        } else {
            System.out.println("转换失败");
        }
    }


    public static boolean wordToPdf(final String wordFile, final String pdfFile) {
        ActiveXComponent app = null;
        try {
            // 打开word
            app = new ActiveXComponent("Word.Application");
            // 设置word不可见
            app.setProperty("Visible", false);
            // 获得word中所有打开的文档
            Dispatch documents = app.getProperty("Documents").toDispatch();
            System.out.println("打开文件: " + wordFile);
            // 打开文档
            Dispatch document = Dispatch.call(documents, "Open", wordFile, false, true).toDispatch();
            // 如果文件存在的话，不会覆盖，会直接报错，所以我们需要判断文件是否存在
            File target = new File(pdfFile);
            if (target.exists()) {
                System.out.println(pdfFile + " 文件已经存在");
                return false;
            }
            System.out.println("另存为: " + pdfFile);
            Dispatch.call(document, "SaveAs", pdfFile, WD_FORMAT_PDF);
            // 关闭文档
            Dispatch.call(document, "Close", false);
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        } finally {
            // 关闭office
            assert app != null;
            app.invoke("Quit", 0);
        }
        return true;
    }

    // /**
    //  * excel转换为pdf
    //  */
    // public static boolean excel2Pdf(String inputFile, String pdfFile) {
    //     try {
    //         ActiveXComponent app = new ActiveXComponent("Excel.Application");
    //         app.setProperty("Visible", false);
    //         Dispatch excels = app.getProperty("Workbooks").toDispatch();
    //         Dispatch excel = Dispatch.call(excels, "Open", inputFile, false,
    //                 true).toDispatch();
    //         Dispatch.call(excel, "ExportAsFixedFormat", XL_FORMAT_PDF, pdfFile);
    //         Dispatch.call(excel, "Close", false);
    //         app.invoke("Quit");
    //         return true;
    //     } catch (Exception e) {
    //         return false;
    //     }
    // }
    //
    // /**
    //  * ppt转换为pdf
    //  */
    // public static boolean ppt2Pdf(String inputFile, String pdfFile) {
    //     try {
    //         ActiveXComponent app = new ActiveXComponent(
    //                 "PowerPoint.Application");
    //         app.setProperty("Visible", false);
    //         Dispatch ppts = app.getProperty("Presentations").toDispatch();
    //
    //         // ReadOnly
    //         // Untitled指定文件是否有标题
    //         // WithWindow指定文件是否可见
    //         Dispatch ppt = Dispatch.call(ppts, "Open", inputFile, true, true, false).toDispatch();
    //         Dispatch.call(ppt, "SaveAs", pdfFile, PPT_FORMAT_PDF);
    //         Dispatch.call(ppt, "Close");
    //         app.invoke("Quit");
    //         return true;
    //     } catch (Exception e) {
    //         return false;
    //     }
    // }
}
