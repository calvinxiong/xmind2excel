# xmind2excel

step 1 使用XMind进行用例设计

必须需遵循以下规则：

1、保证XMind文件只有1个画布、1个根节点。

2、旗帜：功能模块，不同颜色表示重要等级（从前到后重要级别逐步减低）。

3、星星：需求，不同颜色表示重要等级（从前到后重要级别逐步减低）。

4、数字：表示用例和用例等级（1：高、2：中、3：低）。

5、用例上加备注：表示前提条件

6、用例下的第一层分支表示【测试步骤】，第二层分支表示【期望结果】，【测试步骤】和【期望结果】节点禁止使用标识；

    多【测试步骤】和多【期望结果】均使用多个分支表示；

    支持没有【测试步骤】、【期望结果】的用例情况。

7、中心主题（根节点）可以使用备注标识项目的基本信息，但工具转换过程暂不解析。

PS：

a. 旗帜和星星标识的功能模块和需求，在工具解析时都会处理成模块，成为TAP用例模板中的目录；如下图第1个用例的目录是：模块(目录)1-子模块1-子模块1.1

b. XMind中测试用例节点，用例描述一定要包含上层各功能模块节点的关键信息（TAPD中测试计划的执行测试用例时看不到目录信息）。

模板如下：
![输入图片说明](https://images.gitee.com/uploads/images/2021/0517/135149_e9fe359a_9109521.png "1.png")


step 2 使用XMind-Excel转换工具把XMind转换成Excel

在目录下载 xmind2excel-1.1.0-RELEASE.jar 转换工具（需要JDK环境，双击可运行）。

![输入图片说明](https://images.gitee.com/uploads/images/2021/0517/140141_59a3a301_9109521.png "2.png")
![输入图片说明](https://images.gitee.com/uploads/images/2021/0517/140214_c1aef12d_9109521.png "3.png")


step 3 把Excel用例导入到TAPD

![输入图片说明](https://images.gitee.com/uploads/images/2021/0517/140351_3631d130_9109521.png "4.png")
![输入图片说明](https://images.gitee.com/uploads/images/2021/0517/140401_d64ff38c_9109521.png "5.png")


#### 介绍
{**以下是 Gitee 平台说明，您可以替换此简介**
Gitee 是 OSCHINA 推出的基于 Git 的代码托管平台（同时支持 SVN）。专为开发者提供稳定、高效、安全的云端软件开发协作平台
无论是个人、团队、或是企业，都能够用 Gitee 实现代码托管、项目管理、协作开发。企业项目请看 [https://gitee.com/enterprises](https://gitee.com/enterprises)}

#### 软件架构
软件架构说明


#### 安装教程

1.  xxxx
2.  xxxx
3.  xxxx

#### 使用说明

1.  xxxx
2.  xxxx
3.  xxxx

#### 参与贡献

1.  Fork 本仓库
2.  新建 Feat_xxx 分支
3.  提交代码
4.  新建 Pull Request


#### 特技

1.  使用 Readme\_XXX.md 来支持不同的语言，例如 Readme\_en.md, Readme\_zh.md
2.  Gitee 官方博客 [blog.gitee.com](https://blog.gitee.com)
3.  你可以 [https://gitee.com/explore](https://gitee.com/explore) 这个地址来了解 Gitee 上的优秀开源项目
4.  [GVP](https://gitee.com/gvp) 全称是 Gitee 最有价值开源项目，是综合评定出的优秀开源项目
5.  Gitee 官方提供的使用手册 [https://gitee.com/help](https://gitee.com/help)
6.  Gitee 封面人物是一档用来展示 Gitee 会员风采的栏目 [https://gitee.com/gitee-stars/](https://gitee.com/gitee-stars/)
