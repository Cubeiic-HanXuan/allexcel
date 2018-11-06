package com.cubeiic.excel.util;

import java.net.InetAddress;
import java.net.UnknownHostException;

/**
 * @author hanxuan
 * @date 2018/11/6 10:16
 */
public class IpUtil {
    public static void main(String[] args) {
        System.out.println("本机的ip=" + IpUtil.getIp());
    }

    public static String getPorjectPath(){
        String nowpath = "";
        nowpath=System.getProperty("user.dir")+"/";

        return nowpath;
    }

    /**
     * 获取本机ip
     * @return
     */
    public static String getIp(){
        String ip = "";
        try {
            InetAddress inet = InetAddress.getLocalHost();
            ip = inet.getHostAddress();
            //System.out.println("本机的ip=" + ip);
        } catch (UnknownHostException e) {
            e.printStackTrace();
        }

        return ip;
    }
}
