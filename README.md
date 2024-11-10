

## 概述
该脚本用于比较两个 Excel 文件中的 m/z 数据，基于指定的误差范围 tolerance 来查找匹配的行。最终，生成三个输出文件：
	1.	matched.xlsx：包含匹配的 normal 和 cnacer 数据。
	2.	remaining_normal.xlsx：包含 normal 文件中未找到匹配项的数据。
	3.	remaining_cnacer.xlsx：包含 cnacer 文件中未找到匹配项的数据。

## 文件结构

	•	输入文件：
	•	normal.xlsx：包含 normal 数据的 Excel 文件，第一列为 m/z 数据。
	•	cnacer.xlsx：包含 cnacer 数据的 Excel 文件，第一列为 m/z 数据。
	•	输出文件：
	•	matched.xlsx：匹配的 normal 和 cnacer 数据依次排列。
	•	remaining_normal.xlsx：normal 中无匹配的行。
	•	remaining_cnacer.xlsx：cnacer 中无匹配的行。

## 参数说明

	•	tolerance匹配误差范围。如果两个文件中 m/z 值的差绝对值小于或等于该值，则认为该行匹配。

## 环境依赖

```shell
pip install -r requirements.txt
```

## 使用步骤

	1.	将 normal.xlsx 和 cnacer.xlsx 放置在 data 文件夹中。
	2.	在脚本中调整 tolerance 值（如果需要）。
	3.	运行脚本：
```shell
python main.py
```



	4.	脚本运行后，将在 data 文件夹中生成三个输出文件：matched.xlsx、remaining_normal.xlsx、remaining_cnacer.xlsx。

输出文件结构

	•	matched.xlsx：包含 normal 和 cnacer 数据的匹配行。在匹配的每一对数据行中，normal 的列名前缀为 normal_，cnacer 的列名前缀为 cnacer_。
	•	remaining_normal.xlsx：列出 normal-3.xlsx 中未找到匹配的 m/z 数据。
	•	remaining_cnacer.xlsx：列出 cnacer-3.xlsx 中未找到匹配的 m/z 数据。

## 注意事项

	•	输出的匹配文件将根据误差范围查找匹配行，因此误差范围越小，匹配项越少，反之亦然。
	•	如果没有匹配项或没有剩余项，输出的 Excel 文件将生成空表单。

## 示例输出

文件对比完成。生成的文件分别为 matched.xlsx、remaining_normal.xlsx、remaining_cnacer.xlsx

代码结构说明

	•	数据读取：读取 normal-3.xlsx 和 cnacer-3.xlsx，提取第一列 m/z 数据。
	•	匹配查找：逐一对比 normal 和 cnacer 中的 m/z 值，找到误差范围内的匹配项。
	•	结果存储：将匹配项存储到 matched.xlsx，未匹配的 normal 和 cnacer 数据分别存储到 remaining_normal.xlsx 和 remaining_cnacer.xlsx。

常见问题

	1.	没有生成文件：检查 data 文件夹是否存在输入文件。
	2.	匹配结果为空：可能 tolerance 误差范围太小，尝试增大该值。

通过此脚本，您可以高效地比对两个 Excel 文件中的数值，尤其适用于需要匹配误差范围内的研究数据对比。