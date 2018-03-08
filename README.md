# winform-CRUD
这是一个用 WinForm 写的 包含 增删改查 的订单管理系统

## 这是什么
这是一个用 WinForm 写的 包含 增删改查 的订单管理系统

  - 订单查询：实现条件查询，包括种类、型号、颜色、客户名称、订单日期进行查询与统计
  - 订单新增：从数据库读取客户列表，从配置文件中读取种类、型号、颜色参数列表，进行订单插入
  - 订单修改：输入订单号进行修改 
  - 订单删除：输入订单号进行删除
  - 订单导出：导出为 Word 格式，供打印

![订单查询展示](https://www.suiyia.com/wp-content/uploads/2018/03/查询.png)
![订单新增展示](https://www.suiyia.com/wp-content/uploads/2018/03/新增.png)
![订单修改展示](https://www.suiyia.com/wp-content/uploads/2018/03/修改.png)
![订单导出展示](https://www.suiyia.com/wp-content/uploads/2018/03/TIM截图20180308105826.png)


## 使用

  - 连接数据库：MySQL root 123456
  - 系统环境：.NET Framework 4.5

## 引用的第三方库

 - [MaterialSkin 主体界面](https://github.com/IgnaceMaes/MaterialSkin)
 - [NPinyin 汉字转拼音](https://code.google.com/archive/p/npinyin/)
 - [Spire.Doc 操作 Word](http://www.e-iceblue.cn/spiredoc/set-up-new-word-document.html)
 - Mysql.Data 连接数据库
 - System.configuration 读写 App.config 文件实现「热插拔」
 - LogerHelper.cs 日志记录
 - MysqlManager.cs 数据库操作
 - AutoSizeFormClass.cs 可以实现窗口缩放(弃用)
 - public string ConvertSum(string str){}  实现数字转中文大写
 
## 其它 
 - 界面不显示、功能出 bug 要学会看日志文件，在 Debug 目录下
 - 功能并不完善，数据的输入验证还没做
 - 感觉更像一个打印订单的系统，最终统计还是以人工为主
 - 数量可输入负数、用于表示退货......
 - 订单一定要保存，不然存入不了数据库
