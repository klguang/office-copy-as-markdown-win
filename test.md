**Summary**

repository 单独一个maven模块

admin和portal 各自有自己的service、query、convert，复制即可

repository单独一个maven模块admin和portal各自有自己的service、query、convert，复制即可

**复制步骤，必须卸载SonarQube**

**idea不能智能处理红线 import，Ctrl + Shift + R，批量替换**

1. **Maven clean**
2. **将 repository 中按照顺序 controller**、**service、query、convert一个包一个包的拖动到 admin**
3. **必须执行 maven clean install，因为移动service 会出问题，**
4. **将admin中按照顺序**query**、convert、service、controller 复制到portal**

**repository**：是否单独模块，取决于

- 多种数据库支持：
- 数据库变更需求
- Toc系统，同时又有admin

**macro/mall项目参考**

[https://gitee.com/macrozheng/mall](https://gitee.com/macrozheng/mall)

- repository单独模块，mall-mbg（代码生成）
- admin和portal模块各自有自己的
  - DAO包，数据库操作
  - service包，
  - model包（domain），请求参数封装，响应封装
  - Controller包
  - config包

**分层领域模型规约：**

阿里巴巴Java开发手册（黄山版）

- **DO（Data Object**）：此对象与数据库表结构一一对应，通过DAO层向上传输数据源对象。
- **BO（Business Object）**：业务对象，可以由Service层输出的封装业务逻辑的对象。
- Entity：持久化对象，对应数据库中的一条记录
- DTO（Data Transfer Object）：数据传输对象，Service向外传输的对象，即Service返回值，如果Entity满足Service返回值，则不要新建DTO类
- Query：数据查询对象，各层接收上层的查询请求。注意超过 2 个参数的查询封装，禁止使用 Map 类来传输。（controller 参数和service 参数通用）
- VO（View Object）：显示层对象，封装Response返回结果，通常是多个VO的组合体。
- POJO：Entity、DTO、Query、VO都属于POJO
- 重要：如果DTO或者Entity满足Response，不用新建VO类；POJO赋值，必须使用mapstruct，在convert包下，避免手动get、set赋值
- Entity：持久化对象，对应数据库中的一条记录
- VO（Value Object）：值对象，即调用service参数或者返回值
- Query：数据查询对象，各层接收上层的查询请求。注意超过 2 个参数的查询封装，禁止使用 Map 类来传输。（controller 参数和service 参数通用）
- Response：封装Controller返回结果，通常是多个VO的组合体。

重要：如果VO对象满足Controller返回结果，不用新建Response对象；java 值转换必须使用mapstruct，在convert包下，避免手动赋值

| **Entity：**持久化对象，对应数据库中的一条记录 **DTO（Data Transfer Object**）：数据传输对象，Service向外传输的对象，即Service返回值，如果Entity满足Service返回值，则不要新建DTO类 **Query**：数据查询对象，各层接收上层的查询请求。注意超过2个参数的查询封装，禁止使用Map类来传输。（controller参数和service参数通用） VO（View Object）：显示层对象，封装Response返回结果，通常是多个VO的组合体。 **POJO：Entity、DTO**、**Query、**VO都属于POJO **重要：如果DTO或者 Entity 满足 Response，不用新建VO类；POJO 赋值，必须使用mapstruct，在 convert 包下，避免手动 get、set 赋值** |
| --- |

| • Entity：持久化对象，对应数据库中的一条记录 • VO（Value Object）：值对象，即调用service参数或者返回值 • Query：数据查询对象，各层接收上层的查询请求。注意超过2个参数的查询封装，禁止使用Map类来传输。（controller参数和service参数通用） • Response：封装Controller返回结果，通常是多个VO的组合体。 重要：如果VO对象满足Controller返回结果，不用新建Response对象；java值转换必须使用mapstruct，在convert包下，避免手动赋值 |
| --- |

maku-boot的简化模型，从前端传给后端和后端传给前端都是VO一把梭（参数和返回值），**AI 编程 VO 可能更好**