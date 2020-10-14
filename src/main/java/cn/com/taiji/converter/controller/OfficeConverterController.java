package cn.com.taiji.converter.controller;

import cn.com.taiji.converter.jacob.JacobOffice;
import cn.com.taiji.converter.util.FileBase64Tool;
import cn.com.taiji.converter.util.RequestUtil;
import cn.com.taiji.converter.vo.ConvertedTarget;
import com.alibaba.fastjson.JSONObject;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Controller;
import org.springframework.util.FileCopyUtils;
import org.springframework.web.bind.annotation.*;
import sun.misc.BASE64Decoder;

import javax.servlet.http.HttpServletRequest;
import java.io.File;
import java.util.UUID;
import java.util.concurrent.TimeUnit;

/**
 * word 转 pdf
 * 利用jacob调用windows office
 *
 * @author penghongyou
 */
@Controller
@CrossOrigin
@RequestMapping("/api")
@Slf4j
public class OfficeConverterController {

    private static final String TMP_DIR = "C:\\wordTemp";

    @PostMapping("/wordBase64ToPdf")
    @ResponseBody
    public String wordFileToPdf(@RequestBody JSONObject jsonObj, HttpServletRequest request) {
        String fileName = jsonObj.getString("fileName");
        String content = jsonObj.getString("content");
        long st = System.currentTimeMillis();
        System.out.println("\n=====调用转换PDF服务开始:" + RequestUtil.getRemoteIp(request));
        log.info("\n入参fileName:{}", fileName);
        try {
            String sufix = fileName.substring(fileName.lastIndexOf("."));
            String prefix = TMP_DIR + File.separatorChar + UUID.randomUUID().toString();
            File word = new File(prefix + sufix);
            File pdf = new File(prefix + ".pdf");
            byte[] bytes = new BASE64Decoder().decodeBuffer(content);
            FileCopyUtils.copy(bytes, word);

            // 将文件转换为pdf
            ConvertedTarget ct = new ConvertedTarget(word, pdf);
            ct.setBase64File(content);
            JacobOffice.targets.add(ct);
            ct.getCountDownLatch().await(60, TimeUnit.SECONDS);

            // 再将pdf转成Base64
            content = FileBase64Tool.encodeBase64File(ct.getPdf().getAbsolutePath());

            System.out.println(String.format("=====调用转换PDF服务结束,耗时:%s ms", (System.currentTimeMillis() - st)));
            return content;
        } catch (Exception e) {
            System.err.println("文件转换异常");
            e.printStackTrace();
            return null;
        }
    }
}
