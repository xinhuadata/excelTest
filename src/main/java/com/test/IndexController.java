package com.test;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.io.IOUtils;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import com.util.OrderUtil;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;

@Controller
public class IndexController {

	@RequestMapping("/index")
	public String index() {
		return "index";
	}

	@RequestMapping(value = "/upload", method = RequestMethod.POST)
	public String upload(@RequestParam("upload") MultipartFile upload, HttpServletRequest request, Model model) {

		String filename = upload.getOriginalFilename();
		try {
			Workbook workbook = Workbook.getWorkbook(upload.getInputStream());
			Sheet sheet = workbook.getSheet(0);
			String randomTiem = OrderUtil.getOrderIdByTime();
			File file = new File("d:/data/" + randomTiem + ".txt");
			model.addAttribute("filePath", "d:/data/" + randomTiem+ ".txt");
			FileWriter fw = new FileWriter(file);
			BufferedWriter bw = new BufferedWriter(fw);
			// j为行数，getCell("列号","行号")
			int j = sheet.getRows();
			int y = sheet.getColumns();
			for (int i = 0; i < j; i++) {
				for (int x = 0; x < y; x++) {

					Cell c = sheet.getCell(x, i);
					String s = c.getContents();
					bw.write(s);
					bw.write("                          ");
					bw.flush();
				}
				bw.newLine();
				bw.flush();
			}
			System.out.println("写入结束");
		} catch (Exception e) {
			e.printStackTrace();
		}
		return "down";
	}

	@RequestMapping(value ="/fileDownload", method = RequestMethod.POST)
	@ResponseBody
	public String download(HttpServletResponse response,String downLoadPath) throws Exception {
		BufferedInputStream bis = null;
		BufferedOutputStream bos = null;
		String loadPath=downLoadPath;
		// 获取输入流
		bis = new BufferedInputStream(new FileInputStream(downLoadPath));
		String randomTiem = OrderUtil.getOrderIdByTime();
		//获取上传文件的扩展名
		String suffix=downLoadPath.substring(downLoadPath.lastIndexOf("."));
				
		// 输出流
		bos = new BufferedOutputStream(new FileOutputStream(new File("d:/down/"+randomTiem+suffix)));
		byte[] buff = new byte[2048];
		int bytesRead;
		while (-1 != (bytesRead = bis.read(buff, 0, buff.length))) {
			bos.write(buff, 0, bytesRead);
		}
		// 关闭流
		bis.close();
		bos.close();

		return "下载成功";
	}

	
	@RequestMapping(value ="downloadFileAction", method = RequestMethod.POST)
	@ResponseBody
    public void downloadFileAction(HttpServletRequest request, HttpServletResponse response,String downLoadPath) {
 
        response.setCharacterEncoding(request.getCharacterEncoding());
        response.setContentType("application/octet-stream");
        FileInputStream fis = null;
        try {
            File file = new File(downLoadPath);
            fis = new FileInputStream(file);
            response.setHeader("Content-Disposition", "attachment; filename="+file.getName());
            IOUtils.copy(fis,response.getOutputStream());
            response.flushBuffer();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (fis != null) {
                try {
                    fis.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

	}
}
