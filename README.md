# ReflectionExcel

**ReflectionExcel**是一款专为`JavaWeb`打造的Excel数据读取和Excel写入的Java小组件，ReflectionExcel 最大的特点是小巧并且高效，利用Java反射机制原理，从而提高导入和导出效率。
 
- **功能丰富** ：ReflectionExcel 封装了很多转换工具，包括Excel本地读取 ，Excel  流方式读取。
- **使用简单** ：ReflectionExcel 做到开箱即用，不需要任何配置，也不需要任何工具，导入jar包即可使用。
- **使用场景** ：象形一下我们现在有个Excel导入导出功能,比如是导入一个驾驶员花名册,比如驾驶员这个实体类是 `Person.java`  一般的导入是需要每行每列对Excel进行遍历,然后给 `Person` 的各个属性进行赋值,然后如果碰到  `Vehicle.java` 需要导入,难道还要在写一遍遍历?代码冗余不说,而且要写大量工具类对 `POI` 进行Excel版本兼容操作，后期代码维护十分困难。现在借助ReflectionExcel这个小工具,可以实现优雅的Excel读取写入.并且兼容2003，20007两个版本的Excel。

------------------------------------------------------------------------------------------------------------------

## ReflectionExcel使用

### 插件集成
**导入jar到项目：**
```xml
	<dependency>
       <groupId>com.cubeiic</groupId>
       <artifactId>excel</artifactId>
       <version>1.0.0</version>
    </dependency>
```

**ReflectionExcel导入方法调用：**
  - **导入Excel到数据库读取本地文件方式：**
``` java 
// 表对应的实体类 属性 拼接
String keyValue = "从业人员编号:personId,企业标识:companyId,行政区划代码:address,机动车驾驶员姓名:driverName,机动车驾驶员电话:driverPhone,驾驶员性别:driverGender,出生日期:driverBirthday,国籍:driverNationality,驾驶员民族:driverNation,证件编号:credentialsNumber紧急情况联系人:emergencyContact,紧急情况联系人电话:emergencyContactPhone,紧急情况联系人通讯地址:emergencyContactAddress,审核状态:state,创建时间:createTime,更新时间:updateTime";

//调用读取Excel方法（本地方式读取）
List<Person> list =  ReflectionExcel.readXls(file,ReflectionExcel.getMap(keyValue),Person.class.getName());

//批量插入方法 限制一次只能插入20000条
//此方法需要自定义
personService.insertBatch(list,20000);

```
- **导入Excel到数据库读取文件流方式：**
``` java 
// 表对应的实体类 属性 拼接
String keyValue = "从业人员编号:personId,企业标识:companyId,行政区划代码:address,机动车驾驶员姓名:driverName,机动车驾驶员电话:driverPhone,驾驶员性别:driverGender,出生日期:driverBirthday,国籍:driverNationality,驾驶员民族:driverNation,证件编号:credentialsNumber紧急情况联系人:emergencyContact,紧急情况联系人电话:emergencyContactPhone,紧急情况联系人通讯地址:emergencyContactAddress,审核状态:state,创建时间:createTime,更新时间:updateTime";

//调用读取Excel方法（流方式读取）
List<Person> list =  ReflectionExcel.readXls(file.getBytes(),ReflectionExcel.getMap(keyValue),Person.class.getName());

//批量插入方法 限制一次只能插入20000条
//此方法需要自定义
personService.insertBatch(list,20000);

```
**ReflectionExcel导出方法调用：**
- **导出Excel到本地磁盘位置：**
```java
//导入keyValue 拼接字符串
String keyValue = "从业人员编号:personId,企业标识:companyId,行政区划代码:address,机动车驾驶员姓名:driverName,机动车驾驶员电话:driverPhone,驾驶员性别:driverGender,出生日期:driverBirthday,国籍:driverNationality,驾驶员民族:driverNation,证件编号:credentialsNumber紧急情况联系人:emergencyContact,紧急情况联系人电话:emergencyContactPhone,紧急情况联系人通讯地址:emergencyContactAddress,审核状态:state,创建时间:createTime,更新时间:updateTime"

//查询需要导出的数据（此方法需要自定义）
List<Person> personList = personService.selectList(null);

//导出Excel 到本地磁盘
ReflectionExcel.exportExcel(file, keyValue, personList, Person.class.getName(), "从业人员信息");
```
- **在浏览器中直接输出Excel：**
```java
	@RequestMapping(value = "/doExportPerson")
	public ResultEntity doExportPerson(HttpServletResponse response) throws Exception {
		ResultEntity resultEntity = new ResultEntity(ErrorCodeType.SUCCESS, "导出成功", "");
		//导入keyValue 拼接字符串
		String keyValue = "从业人员编号:personId,企业标识:companyId,行政区划代码:address,机动车驾驶员姓名:driverName,机动车驾驶员电话:driverPhone,驾驶员性别:driverGender,出生日期:driverBirthday,国籍:driverNationality,驾驶员民族:driverNation,证件编号:credentialsNumber紧急情况联系人:emergencyContact,紧急情况联系人电话:emergencyContactPhone,紧急情况联系人通讯地址:emergencyContactAddress,审核状态:state,创建时间:createTime,更新时间:updateTime"
		try {
			List<Person> personList = personService.selectList(null);
			if (personList.size() != 0){
				//在浏览器中直接输出 exportExcel
				ReflectionExcel.exportExcelOutputStream(response, keyValue, personList, Person.class.getName(), "从业人员信息");
			}else {
				resultEntity = new ResultEntity(ErrorCodeType.SUCCESS, "导出数据不能为空", "");
			}

		} catch (Exception e) {
			e.printStackTrace();
			resultEntity = new ResultEntity(ErrorCodeType.SUCCESS, "导出数据异常", "");
		}
		return resultEntity;
	}
```

## 反馈与建议
- QQ：768519234
- 邮箱：<chendebingem@163.com>

---------
感谢阅读这份帮助文档。请点击右上角，开启全新的记录与分享体验吧。
