package cn.com.taiji.converter.vo;

import lombok.Data;
import lombok.experimental.Accessors;

import java.io.File;
import java.util.concurrent.CountDownLatch;

/**
 * @author penghongyou
 */
@Data
@Accessors
public class ConvertedTarget {

    //标识该文件已被处理,并通知线程
    private CountDownLatch countDownLatch = new CountDownLatch(1);
    private File word;
    private File pdf;
    private String base64File;
    private byte[] byteFile;
    private int tryCount;

    public ConvertedTarget(File word, File pdf) {
        this.word = word;
        this.pdf = pdf;
    }
}