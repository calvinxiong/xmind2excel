package cn.xxx.xmind2excel.util;

import org.apache.commons.io.FileUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.nio.charset.Charset;
import java.util.Enumeration;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

/**
 * @author xiongchenghui
 * @date 2020-08-14
 * &Desc 文件工具类
 */
public class FileUtil {
    private static Logger logger = LoggerFactory.getLogger(FileUtil.class);

    /***
     * &Desc: xMind文件转换为zip文件
     * @param xMindFile xMind文件
     * @return java.io.File
     */
    public static File transferXMind2Zip(File xMindFile) {
        File file = new File("Extract"+ File.separator + xMindFile.getName());

        // 如果Extract目录不存在, 则创建
        File pFile = new File(file.getParent());
        if (!pFile.exists()) {
            pFile.mkdir();
        }

        // 将文件复制一份
        FileUtil.copyFileUsingApacheCommonsIO(xMindFile, file);
        logger.info("使用备份xMind文件转换,路径为" + file.getAbsolutePath());

        // 获取即将转换为zip的文件的绝对路径 \\转译符
        String zipFileDest = file.getAbsolutePath().replaceAll("\\" + FileExtension.XMIND, FileExtension.ZIP);
        File zipFile = new File(zipFileDest);
        // 如果zip文件存在先删除
        if(zipFile.exists()){
            zipFile.delete();
        }
        // 将xMind文件转换为zip文件
        file.renameTo(zipFile);
        logger.info("转换后的zip文件路径为:" + zipFileDest);

        return zipFile;
    }

    /***
     * &Desc: 解压文件到指定目录
     * @param zipPath 解压文件路径
     * @param descDir 解压到指定目录
     * @return void
     */
    @SuppressWarnings("rawtypes")
    public static void unZipFiles(String zipPath, String descDir) throws IOException {
        File zipFile = new File(zipPath);
        File pathFile = new File(descDir);

        logger.info("******************开始解压********************");

        // 检查解压目录是否存在
        if(pathFile.exists()){
            deleteDir(pathFile);
        }
        pathFile.mkdirs();

        // 解决zip文件中有中文目录或者中文文件
        ZipFile zip = new ZipFile(zipFile, Charset.forName("GBK"));
        for (Enumeration entries = zip.entries(); entries.hasMoreElements();) {
            ZipEntry entry = (ZipEntry) entries.nextElement();
            String zipEntryName = entry.getName();
            InputStream in = zip.getInputStream(entry);
            String outPath = (descDir + "/" + zipEntryName);

            // 判断路径是否存在,不存在则创建文件路径
            File file = new File(outPath.substring(0, outPath.lastIndexOf('/')));
            if (!file.exists()) {
                file.mkdirs();
            }

            // 判断文件全路径是否为文件夹,如果是, 则不需要解压
            if (new File(outPath).isDirectory()) {
                continue;
            }

            // 执行解压
            OutputStream out = new FileOutputStream(outPath);
            byte[] buf = new byte[1024];
            int len;
            while ((len = in.read(buf)) > 0) {
                out.write(buf, 0, len);
            }

            // 关闭文件流
            in.close();
            out.close();
        }
        zip.close();
        logger.info("******************解压完毕********************");
    }

    /***
     * &Desc: 文件复制
     * @param source 源文件
     * @param dest 目标文件
     * @return void
     */
    private static void copyFileUsingApacheCommonsIO(File source, File dest) {
        try {
            FileUtils.copyFile(source, dest);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /***
     * &Desc: 删除带文件的目录
     * @param dir 目录
     * @return boolean
     */
    public static boolean deleteDir(File dir) {
        //先删除目录下的所有文件
        if (dir.isDirectory()) {
            String[] children = dir.list();

            //递归删除目录中的子目录下
            for (int item=0; item<children.length; item++) {
                boolean success = deleteDir(new File(dir, children[item]));
                if (!success) {
                    System.out.println("递归删除文件失败："+ new File(dir, children[item]).getAbsolutePath() );
                    return false;
                }
            }
        }

        // 目录此时为空，可以删除
        return dir.delete();

    }


}