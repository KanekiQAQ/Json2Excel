# Json2Excel

一个把 Json 转换成 Excel 的工具。

目前只支持一个层级。

```JSON
[
    {
        "name": "test1",
        "age": 11
    },
    {
        "name": "test2",
        "age": 12
    },
    {
        "name": "test3"
    }
]
```
转换后
|name	|age|
|-----|-----|
|test1	|11|
|test2	|12|
|test3	||
