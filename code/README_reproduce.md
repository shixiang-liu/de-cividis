# DE-Cividis 复现说明

当前论文以 `论文/计（创）2302班-刘世翔-XXXXXXX-数字图像处理技术结课论文.docx` 和同名 PDF 为准。脚本用于刷新实验结果、正文图片、Word 排版和 PDF 导出。

## 一键刷新

在 `Essay` 目录运行：

```powershell
python code/write_paper_decividis_final.py
```

如需重新计算全部实验指标，再运行：

```powershell
python code/write_paper_decividis_final.py --with-experiments
```

全量实验使用 51 幅图像：主评估集 27 幅，Kodak 补充验证集 24 幅。完整运行会写入 `results/final_results.csv` 和 `results/final_results_summary.csv`，再生成图片、刷新 DOCX，并通过 LibreOffice 导出 PDF。

## 正文图片

- 图1：`figs/fig1_pipeline.png`，DE-Cividis 处理流程。
- 图2：`figs/fig2_colormap_cvd.png`，不同色图在色觉障碍模拟下的对比。
- 图3：`figs/fig3_saliency.png`，参数对输出效果的影响。
- 图4：`figs/fig4_eval_overview.png`，27 幅主评估图像的 DE-Cividis 输出总览。
- 图5：`figs/fig5_typical_cases.png`，X 射线、焊缝和纹理图像的典型对比。
- 图6：`figs/fig6_ablation.png`，DE-Cividis 消融实验。
- 图7：`figs/fig7_kodak_overview.png`，Kodak 01-24 补充验证输出总览。
