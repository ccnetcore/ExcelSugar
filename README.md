<h4 align="center">一套以用户体验出发的.netstandard2.1 Excel开源工具包</h4>
<h5 align="center">支持查询、导出、模板等常用功能,语法与Sqlsugar类似</h5>
<h2 align="center">Excel快速上手的不二之选</h2>

#### ExcelSugar 介绍

ORM To Excel,像操作对象一样操作Excel，支持查询、导出、模板等常用功能;

使用方式类比于SqlsugarORM，优雅简单，excel快速上手的不二之选;

具备的完整的单元测试，可以在仓库中查看及详细使用;

> 正在持续迭代中，支持的功能会越来越多~~！
#### 优势

1. [x] 将Excel当作`对象`操作，使用简单，1分钟上手
2. [x] 语法与`SqlsugarORM`基本类似，两者都可以兼得
3. [x] 支持`where`表达式查询
4. [x] 支持来自`模板`导出
5. [x] 支持`动态表头`功能
6. [x] 支持单元格存储`对象`
7. [x] 完善的`单元测试`，可以查看详细使用操作
8. [x] 作者`24`小时在线，可QQ联系：454313500

#### 安装教程

Nuget包管理器直接搜索ExcelSugar.Npoi安装即可（后续可更换实现）
- ExcelSugar.Core（核心包-自动依赖）
- ExcelSugar.Npoi（实现包，下载）
- ExcelSugar.AspNetCore(正在开发)

#### 使用说明

1.  创建ExcelClient
``` cs
IExcelSugarClient excelSugarClient = new ExcelSugarClient(new ExcelSugarConfig { Path = "../../../TempExcel/Test.xlsx", HandlerType = ExcelHandlerType.Npoi });

```
2.  定义模型，并打上对应的特性
``` cs
 [SugarSheet("测试表")]
 public class TestModel
 {
     [SugarHead("姓名")]
     public string Name { get; set; }
     [SugarHead("描述")]
     public string Description { get; set; }

     /// <summary>
     /// 导出可支持动态列模型
     /// </summary>
     [SugarDynamicHead]
     public List<DynamicModel> DynamicModels { get; set; } = new List<DynamicModel>();

     //正在支持中。。。。。。理论上支持
     //[SugarHead("对象列", IsJson = true)]
     //public Dictionary<int, string> KKK { get; set; } = new Dictionary<int, string>();
 }
 /// <summary>
 /// 动态表头，需具备IsCode、IsValue、IsName 3个字段
 /// </summary>
 public class DynamicModel
 {
     [SugarDynamicHead(IsCode = true)]
     public string DataCode { get; set; }

     [SugarDynamicHead(IsValue = true)]
     public int DataValue { get; set; }

     [SugarDynamicHead(IsName = true)]
     public string DataName { get; set; }
 }

```
3. 使用client操作即可
``` cs
//创建实体
var testModel = new List<TestModel> {
	new TestModel { Description = "男的", Name = "张三", DynamicModels=new List<DynamicModel>{
		new DynamicModel { DataCode="height",DataName="身高",DataValue=188},
	} },
	new TestModel { Description = "女的", Name = "李四" ,DynamicModels=new List<DynamicModel>{
		new DynamicModel { DataCode="height",DataName="身高",DataValue=168},
		new DynamicModel { DataCode="age",DataName="年龄",DataValue=18},
	} }
};

//从excel中查询出实体
var data = await client.Queryable<TestModel>().ToListAsync();

//从excel中查询出实体，支持where表达式
var data = await client.Queryable<TestModel>().Where(x => x.Name == "张三").ToListAsync();

//将实体直接导出excel
await client.Exportable(testModel).ExecuteCommandAsync();

//将实体根据给定的模板进行导出excel
await client.Exportable(testModel).From("../../../TempExcel/Template.xlsx").ExecuteCommandAsync();
//Or await client.Exportable("../../../Test.xlsx").ExecuteCommandAsync();

```

#### 致谢

- [Sqlsugar老杰哥] https://www.donet5.com/Home/Doc


#### 联系我们

- 可通过QQ进行联系：454313500
- 可通过本仓库进行联系
