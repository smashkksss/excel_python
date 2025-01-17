import json
import openpyxl
from openpyxl import Workbook

# JSON 文件路径
json_file = "data01.json"

try:
    with open(json_file, "r", encoding="utf-8") as file:
        data = json.load(file)  # 尝试解析
        print("JSON data loaded！")
        print(data)
except json.JSONDecodeError as e:
    print(f"JSONDecodeError: {e}")
    print("check JSONfile ！")

# 创建工作簿
wb = Workbook()
ws = wb.active
ws.title = "Betting Data"

# 添加标题行
headers = [
    "Sport Key",
    "Sport Title",
    "Commence Time",
    "Home Team",
    "Away Team",
    "Bookmaker Key",
    "Bookmaker Title",
    "Market Key",
    "Outcome Team",
    "Odds",
]
ws.append(headers)

# 解析数据并写入 Excel
for event in data["data"]:
    sport_key = event["sport_key"]
    sport_title = event["sport_title"]
    commence_time = event["commence_time"]
    home_team = event["home_team"]
    away_team = event["away_team"]

    for bookmaker in event["bookmakers"]:
        bookmaker_key = bookmaker["key"]
        bookmaker_title = bookmaker["title"]

        for market in bookmaker["markets"]:
            market_key = market["key"]

            for outcome in market["outcomes"]:
                outcome_team = outcome["name"]
                odds = outcome["price"]

                # 写入一行数据
                ws.append([
                    sport_key,
                    sport_title,
                    commence_time,
                    home_team,
                    away_team,
                    bookmaker_key,
                    bookmaker_title,
                    market_key,
                    outcome_team,
                    odds,
                ])

# 保存到 Excel 文件
output_file = "betting_data.xlsx"
wb.save(output_file)

print(f"data file  {json_file} Into :{output_file}")

