package xyz.baal.pdf;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;
import com.lowagie.text.Document;
import com.lowagie.text.DocumentException;
import com.lowagie.text.Image;
import com.lowagie.text.PageSize;
import com.lowagie.text.pdf.PdfWriter;

/**
 * 功能： 将Office文件转换为PDF文档
 * 
 * 依赖外部jar包：jacob.jar(包括jacob-1.18-x64.dll，此文件要放在C:\Windows\System32\下)
 * com.lowagie.text-2.1.7.jar
 * 
 * @author
 *
 */
public class FileToPDF {

	static final int wdDoNotSaveChanges = 0;// 不保存待定的更改。
	static final int wdFormatPDF = 17;// word转PDF 格式
	static final int ppSaveAsPDF = 32;// ppt 转PDF 格式

	public static void main(String[] args) throws IOException {
		String source = "d:\\word.doc";
		String topdf = "d:\\test1.pdf";

		FileToPDF pdf = new FileToPDF();
		pdf.WordToPDF(source, topdf);
	}

	public void WordToPDF(String source, String target) {

		long start = System.currentTimeMillis();
		ActiveXComponent app = null;
		try {
			app = new ActiveXComponent("Word.Application");
			app.setProperty("Visible", false);

			Dispatch docs = app.getProperty("Documents").toDispatch();
			System.out.println("打开文档" + source);
			Dispatch doc = Dispatch.call(docs, "Open", source, false, true).toDispatch();

			System.out.println("转换文档到PDF " + target);
			File tofile = new File(target);
			if (tofile.exists()) {
				tofile.delete();
			}
			Dispatch.call(doc, "SaveAs", target, wdFormatPDF);

			Dispatch.call(doc, "Close", false);
			long end = System.currentTimeMillis();
			System.out.println("转换完成..用时：" + (end - start) + "ms.");
		} catch (Exception e) {
			System.out.println("========Error:文档转换失败：" + e.getMessage());
		} finally {
			if (app != null)
				app.invoke("Quit", wdDoNotSaveChanges);
		}
	}

	public void PptToPDF(String source, String target) {
		long start = System.currentTimeMillis();
		ActiveXComponent app = null;
		try {
			app = new ActiveXComponent("Powerpoint.Application");
			Dispatch presentations = app.getProperty("Presentations").toDispatch();
			System.out.println("打开文档" + source);
			Dispatch presentation = Dispatch.call(presentations, "Open", source, true, true, false).toDispatch();

			System.out.println("转换文档到PDF " + target);
			File tofile = new File(target);
			if (tofile.exists()) {
				tofile.delete();
			}
			Dispatch.call(presentation, "SaveAs", target, ppSaveAsPDF);

			Dispatch.call(presentation, "Close");
			long end = System.currentTimeMillis();
			System.out.println("转换完成..用时：" + (end - start) + "ms.");
		} catch (Exception e) {
			System.out.println("========Error:文档转换失败：" + e.getMessage());
		} finally {
			if (app != null)
				app.invoke("Quit");
		}
	}

	public void ExcelToPDF(String source, String target) {
		long start = System.currentTimeMillis();
		ActiveXComponent app = new ActiveXComponent("Excel.Application");
		try {
			app.setProperty("Visible", false);
			Dispatch workbooks = app.getProperty("Workbooks").toDispatch();
			System.out.println("打开文档" + source);
			Dispatch workbook = Dispatch.invoke(workbooks, "Open", Dispatch.Method,
					new Object[] { source, new Variant(false), new Variant(false) }, new int[3]).toDispatch();
			Dispatch.invoke(workbook, "SaveAs", Dispatch.Method,
					new Object[] { target, new Variant(57), new Variant(false), new Variant(57), new Variant(57),
							new Variant(false), new Variant(true), new Variant(57), new Variant(true),
							new Variant(true), new Variant(true) },
					new int[1]);
			Variant f = new Variant(false);
			System.out.println("转换文档到PDF " + target);
			Dispatch.call(workbook, "Close", f);
			long end = System.currentTimeMillis();
			System.out.println("转换完成..用时：" + (end - start) + "ms.");
		} catch (Exception e) {
			System.out.println("========Error:文档转换失败：" + e.getMessage());
		} finally {
			if (app != null) {
				app.invoke("Quit", new Variant[] {});
			}
		}
	}

	public boolean ImgToPDF(String imgFilePath, String pdfFilePath) throws IOException {
		File file = new File(imgFilePath);
		if (file.exists()) {
			Document document = new Document();
			FileOutputStream fos = null;
			try {
				fos = new FileOutputStream(pdfFilePath);
				PdfWriter.getInstance(document, fos);

				// 添加PDF文档的某些信息，比如作者，主题等等
				document.addAuthor("root");
				document.addSubject("file to pdf.");
				// 设置文档的大小
				document.setPageSize(PageSize.A4);
				// 打开文档
				document.open();
				// 写入一段文字
				// document.add(new Paragraph("JUST TEST ..."));
				// 读取一个图片
				Image image = Image.getInstance(imgFilePath);
				float imageHeight = image.getScaledHeight();
				float imageWidth = image.getScaledWidth();
				int i = 0;
				while (imageHeight > 500 || imageWidth > 500) {
					image.scalePercent(100 - i);
					i++;
					imageHeight = image.getScaledHeight();
					imageWidth = image.getScaledWidth();
					System.out.println("imageHeight->" + imageHeight);
					System.out.println("imageWidth->" + imageWidth);
				}

				image.setAlignment(Image.ALIGN_CENTER);
				// //设置图片的绝对位置
				// image.setAbsolutePosition(0, 0);
				// image.scaleAbsolute(500, 400);
				// 插入一个图片
				document.add(image);
			} catch (DocumentException de) {
				System.out.println(de.getMessage());
			} catch (IOException ioe) {
				System.out.println(ioe.getMessage());
			}
			document.close();
			fos.flush();
			fos.close();
			return true;
		} else {
			return false;
		}
	}
}