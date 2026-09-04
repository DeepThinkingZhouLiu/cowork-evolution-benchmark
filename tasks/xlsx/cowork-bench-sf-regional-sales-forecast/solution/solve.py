#!/usr/bin/env python3
"""根据静态销售和 GDP 快照生成单一区域预测工作簿。"""
from __future__ import annotations
import argparse, json
from collections import defaultdict
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.table import Table, TableStyleInfo

REGIONS=["Азиатско-Тихоокеанский регион","Европа","Латинская Америка","Ближний Восток","Северная Америка"]

def load(path): return json.loads(path.read_text(encoding="utf-8"))
def style(sheet,widths,name):
    for cell in sheet[1]: cell.fill=PatternFill("solid",fgColor="17365D"); cell.font=Font(name="Arial",bold=True,color="FFFFFF"); cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    sheet.freeze_panes="A2"; sheet.auto_filter.ref=sheet.dimensions; sheet.sheet_view.showGridLines=False
    for col,width in widths.items(): sheet.column_dimensions[col].width=width
    table=Table(displayName=name,ref=sheet.dimensions); table.tableStyleInfo=TableStyleInfo(name="TableStyleMedium2",showRowStripes=True); sheet.add_table(table)

def calculate(input_dir):
    historical=load(input_dir/"historical_sales.json"); economics=load(input_dir/"economic_indicators.json")["regions"]; groups=defaultdict(list)
    for row in historical: groups[row["region"]].append(float(row["revenue"]))
    forecasts={}
    for region in REGIONS:
        raw=sum(groups[region])/3; trailing=round(raw,2); gdp=float(economics[region]["gdp_growth"]); adjustment=round(1+gdp/100,4); monthly=round(trailing*adjustment,2); quarterly=round(monthly*3,2)
        forecasts[region]={"trailing":trailing,"gdp":gdp,"adjustment":adjustment,"monthly":monthly,"quarterly":quarterly,"unrounded":raw*adjustment*3}
    total_q4=round(sum(float(row["revenue"]) for row in historical),2); total_q1=round(sum(round(item["unrounded"],2) for item in forecasts.values()),2); high=max(forecasts,key=lambda r:forecasts[r]["unrounded"]); low=min(forecasts,key=lambda r:forecasts[r]["unrounded"]); high_gdp=max(forecasts,key=lambda r:forecasts[r]["gdp"])
    summary={"Total_Q4_2025_Revenue":total_q4,"Total_Monthly_Average":round(total_q4/3,2),"Total_Q1_2026_Forecast":total_q1,"QoQ_Growth_Pct":round((total_q1/total_q4-1)*100,1),"Highest_Forecast_Region":high,"Highest_Forecast_Amount":round(forecasts[high]["unrounded"],2),"Lowest_Forecast_Region":low,"Lowest_Forecast_Amount":round(forecasts[low]["unrounded"],2),"Highest_GDP_Growth_Region":high_gdp,"Regions_Count":len(forecasts)}
    return historical,forecasts,summary

def build(input_dir,output):
    historical,forecasts,summary=calculate(input_dir); wb=Workbook(); wb.properties.title="Regional Sales Forecast"; wb.calculation.fullCalcOnLoad=True; wb.calculation.forceFullCalc=True; wb.calculation.calcMode="auto"
    source=wb.active; source.title="Historical"; source.append(["Region","Month","Orders","Revenue"]); lookup={(x["region"],x["month"]):x for x in historical}
    for r in REGIONS:
        for month in ("2025-10","2025-11","2025-12"):
            row=lookup[(r,month)]; source.append([r,month,int(row["orders"]),round(float(row["revenue"]),2)])
    for cell in source["D"][1:]: cell.number_format='$#,##0.00'
    style(source,{"A":34,"B":14,"C":12,"D":18},"Historical_Sales")
    forecast=wb.create_sheet("Forecast"); forecast.append(["Region","Trailing_3M_Avg","GDP_Growth_Pct","Economic_Adjustment","Adjusted_Monthly_Forecast","Next_Quarter_Forecast"])
    for r in REGIONS:
        f=forecasts[r]; forecast.append([r,f["trailing"],f["gdp"],f["adjustment"],f["monthly"],f["quarterly"]])
    for col in ("B","E","F"):
        for cell in forecast[col][1:]: cell.number_format='$#,##0.00'
    for cell in forecast["C"][1:]: cell.number_format='0.0"%"'
    for cell in forecast["D"][1:]: cell.number_format='0.0000'
    style(forecast,{"A":34,"B":20,"C":18,"D":21,"E":28,"F":24},"Regional_Forecast")
    summary_sheet=wb.create_sheet("Summary"); summary_sheet.append(["Metric","Value"])
    for key,value in summary.items(): summary_sheet.append([key,value])
    for cell in summary_sheet["B"][1:]:
        if isinstance(cell.value,(int,float)): cell.number_format='$#,##0.00' if cell.row not in (5,11) else '0.0"%"'
    style(summary_sheet,{"A":32,"B":38},"Forecast_Summary")
    notes=wb.create_sheet("Forecast Notes"); notes.append(["Раздел","Содержание"])
    notes.append(["Краткое резюме",f"Прогноз выручки на I квартал 2026 года: ${summary['Total_Q1_2026_Forecast']:,.2f}; IV квартал 2025 года: ${summary['Total_Q4_2025_Revenue']:,.2f}; рост: {summary['QoQ_Growth_Pct']:.1f}%."])
    notes.append(["Методология","Для каждого региона: средняя выручка за октябрь-декабрь 2025 × (1 + рост ВВП / 100); квартальный прогноз равен трём месячным."])
    notes.append(["Региональный анализ","; ".join(f"{r}: ${forecasts[r]['quarterly']:,.2f}, ВВП {forecasts[r]['gdp']:.1f}%" for r in REGIONS)])
    notes.append(["Ключевые допущения","Предполагаются стабильные цены, продуктовый портфель, охват команды продаж и экономические условия; новые промоакции и рыночные шоки не моделируются."])
    notes.append(["Рекомендации",f"Усилить коммерческий фокус на {summary['Highest_Forecast_Region']}; для Европы использовать целевые кампании и ежемесячно сверять прогноз с фактом."])
    style(notes,{"A":26,"B":120},"Forecast_Notes")
    for row in notes.iter_rows(min_row=2):
        for cell in row: cell.alignment=Alignment(vertical="top",wrap_text=True)
    output.mkdir(parents=True,exist_ok=True); wb.save(output/"Sales_Forecast.xlsx"); wb.close()

def main():
    ap=argparse.ArgumentParser(); ap.add_argument("--input-dir",type=Path,default=Path("/app/data/input_files")); ap.add_argument("--output-dir",type=Path,default=Path("/app/output")); args=ap.parse_args(); build(args.input_dir,args.output_dir)
if __name__=="__main__": main()
