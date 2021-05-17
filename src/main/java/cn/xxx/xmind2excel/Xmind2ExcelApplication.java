package cn.xxx.xmind2excel;

import cn.xxx.xmind2excel.ui.SwingClient;
import org.jb2011.lnf.beautyeye.BeautyEyeLNFHelper;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.builder.SpringApplicationBuilder;
import org.springframework.context.ApplicationContext;

import javax.swing.*;

@SpringBootApplication
public class Xmind2ExcelApplication {
    public Xmind2ExcelApplication(){
        SwingClient.getInstance().initUI();
    }

    public static void main(String[] args) {
        /*SpringApplication.run(Xmind2ExcelApplication.class, args);*/
        try {
            org.jb2011.lnf.beautyeye.BeautyEyeLNFHelper.launchBeautyEyeLNF();
            BeautyEyeLNFHelper.frameBorderStyle = BeautyEyeLNFHelper.FrameBorderStyle.translucencyAppleLike;
            UIManager.put("RootPane.setupButtonVisible", false);
        }catch (Exception e){

        }
        ApplicationContext ctx = new SpringApplicationBuilder(Xmind2ExcelApplication.class)
                .headless(false).run(args);

    }

}
