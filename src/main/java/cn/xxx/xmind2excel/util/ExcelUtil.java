package cn.xxx.xmind2excel.util;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;


/**
 * @author xiongchenghui
 * @date 2020-08-14
 * &Desc Excel工具类
 */
public class ExcelUtil {

    /***
     * &Desc: 读取Excel文件到Workbook
     * @param filePath 读取文件的路径
     * @return org.apache.poi.ss.usermodel.Workbook
     */
    @SuppressWarnings("resource")
    public static Workbook readExcel(String filePath) {
        Workbook wb = null;
        if (filePath == null) {
            return null;
        }
        String extString = filePath.substring(filePath.lastIndexOf("."));
        InputStream is = null;
        try {
            is = new FileInputStream(filePath);
            if (FileExtension.XLS.equals(extString)) {
                wb = new HSSFWorkbook(is);
            } else if (FileExtension.XLSX.equals(extString)) {
                wb = new XSSFWorkbook(is);
            } else {
                wb = null;
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }finally {
            try {
                is.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return wb;
    }

    /***
     * &Desc: 创建Excel文件
     * @param fileName 创建文件的路径
     * @return void
     */
    public static void createExcel(String fileName) {
        Workbook wb = null;
        if (fileName.endsWith(FileExtension.XLS)) {
            wb = new HSSFWorkbook();
            wb.createSheet();
        }
        if (fileName.endsWith(FileExtension.XLSX)) {
            wb = new XSSFWorkbook();
            wb.createSheet();
        }
        File file = new File(fileName);
        if(file.exists()){
            file.delete();
        }
        OutputStream fileOut = null;
        try {
            fileOut = new FileOutputStream(fileName);
            wb.write(fileOut);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                fileOut.close();
                wb.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }


}