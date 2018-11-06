package com.cubeiic.excel.util;

import org.springframework.context.ApplicationContext;

/**
 * 资源存放地址定义
 * @author hanxuan
 * @date 2017.04.24
 * @version v1.0.0
 * 
 */

public class Const {
	public static final String SESSION_SECURITY_CODE = "sessionSecCode";
	public static final String SESSION_USER = "sessionUser";
	public static final String SESSION_ROLE_RIGHTS = "sessionRoleRights";
	/**
	 * 当前菜单
	 */
	public static final String SESSION_menuList = "menuList";
	/**
	 * 全部菜单
	 */
	public static final String SESSION_allmenuList = "allmenuList";
	public static final String SESSION_QX = "QX";
	public static final String SESSION_userpds = "userpds";
	/**
	 * 用户对象
	 */
	public static final String SESSION_USERROL = "USERROL";
	/**
	 * 用户名
	 */
	public static final String SESSION_USERNAME = "USERNAME";
	public static final String TRUE = "T";
	public static final String FALSE = "F";
	/**
	 * 系统名称路径
	 */
	public static final String LOGIN = "/login_toLogin.do";
	/**
	 * 系统名称路径
	 */
	public static final String SYSNAME = "admin/config/SYSNAME.txt";
	public static final String PAGE	= "admin/config/PAGE.txt";
	/**
	 * 邮箱服务器配置路径
	 */
	public static final String EMAIL = "admin/config/EMAIL.txt";
	/**
	 * 短信账户配置路径1
	 */
	public static final String SMS1 = "admin/config/SMS1.txt";
	/**
	 * 短信账户配置路径2
	 */
	public static final String SMS2 = "admin/config/SMS2.txt";
	/**
	 * 文字水印配置路径
	 */
	public static final String FWATERM = "admin/config/FWATERM.txt";
	/**
	 * 图片水印配置路径
	 */
	public static final String IWATERM = "admin/config/IWATERM.txt";
	/**
	 * 微信配置路径
	 */
	public static final String WEIXIN	= "admin/config/WEIXIN.txt";
	/**
	 * WEBSOCKET配置路径
	 */
	public static final String WEBSOCKET = "admin/config/WEBSOCKET.txt";
	/**
	 * 图片上传路径
	 */
	public static final String FILEPATHIMG = "uploadFiles/uploadImgs/";
	/**
	 * 文件上传路径
	 */
	public static final String FILEPATHFILE = "uploadFiles/file/";
	/**
	 * 二维码存放路径
	 */
	public static final String FILEPATHTWODIMENSIONCODE = "uploadFiles/twoDimensionCode/";
	/**
	 * 不对匹配该值的访问路径拦截（正则）
	 */
	public static final String NO_INTERCEPTOR_PATH = ".*/((login)|(logout)|(code)|(app)|(weixin)|(static)|(main)|(websocket)).*";
	/**
	 * 每页显示最大数100条
	 * 配置分页查询每页最大数
	 */
	public static final Integer MAX_PAGE_SIZE = 100;

	/**
	 * 该值会在web容器启动时由WebAppContextListener初始化
	 */
	public static ApplicationContext WEB_APP_CONTEXT = null;

	
	/**
	 * 首页学员数据统计标识
	 * @desc 
	 * @author jiangxl
	 */
	public interface StuStatiscFlag {
		String YEAR_STU_FLAG = "yearStu";
		String YEAR_GRADUATE_STU_FLAG = "yearGraduateStu";
		String MONTH_STU_FLAG = "monthStu";
		String MONTH_GRADUATE_STU_FLAG = "monthGraduateStu";
		String NOW_STU_FLAG = "nowStu";
		String NOW_GRADUATE_STU_FLAG = "nowGraduateStu";

	}

}
