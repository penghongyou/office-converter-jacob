package cn.com.taiji.converter.util;

import org.springframework.util.ObjectUtils;

import javax.servlet.http.HttpServletRequest;

public class RequestUtil {

    /**
     * 获取ip地址
     *
     * @param request
     * @return
     */
    public static String getRemoteIp(HttpServletRequest request) {
        if (ObjectUtils.isEmpty(request)) {
            return null;
        }
        String forwardIp = request.getHeader("x-forwarded-for");
        String realIp = request.getHeader("X-Real-IP");
        String remoteIp = request.getRemoteAddr();

        if (!ObjectUtils.isEmpty(forwardIp)) {
            return forwardIp.split(",")[0];
        } else if (!ObjectUtils.isEmpty(realIp)) {
            return realIp;
        } else {
            return remoteIp;
        }
    }
}
