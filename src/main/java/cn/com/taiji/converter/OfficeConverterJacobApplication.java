package cn.com.taiji.converter;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

/**
 * windows环境下利用office转换pdf
 * 支持签章 office360签章插件
 * 使用jacob调用office api
 *
 * @author penghongyou
 */
@SpringBootApplication
public class OfficeConverterJacobApplication {

    public static void main(String[] args) {
        SpringApplication.run(OfficeConverterJacobApplication.class, args);
    }

}
