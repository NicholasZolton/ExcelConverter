import xlwings as xw
from datetime import datetime


def main():
    itemResults = {}
    idToName = {}
    nameToId = {}
    indexToId = {}
    indexToName = {}
    with xw.Book("./data/data.xlsx", mode="r") as book:
        # figure out the item ids
        firstWeek = book.sheets[4]
        # print(firstWeek.name)

        # parse item ids
        itemIds = list(firstWeek.range("b4:c12").value)
        for i, (itemId, itemName) in enumerate(itemIds):
            idToName[int(itemId)] = itemName
            indexToId[i] = itemId
            indexToName[i] = itemName
            nameToId[itemName] = itemId

        # each item result is a list of the following:
        # {item_name: [[date, order (full), order (short)]]}

        sheets = list(book.sheets)[4:]
        for sheet in sheets:
            # print(sheet.name)

            # get D3:O12
            data = sheet.range("d3:o12").value
            # print(data)

            # map the first row (dates) to "m/d/y"
            dates = data[0]
            for i, date in enumerate(dates):
                dates[i] = date.strftime("%m/%d/%Y")
            # print(dates)

            # if col % 2 == 1: then the row is a data row
            for row in range(1, len(data)):
                # dataRow is {date: [order (full), order (short)]}
                dataRow = {}
                for col in range(0, len(data[row])):
                    # print(f"data row: {data[row]}")
                    if col % 2 == 1:
                        # short
                        dataRow[dates[col]] = dataRow.get(dates[col], []) + [
                            data[row][col]
                        ]
                    elif col % 2 == 0:
                        # full
                        dataRow[dates[col]] = dataRow.get(dates[col], []) + [
                            data[row][col]
                        ]
                # get the item name
                itemName = indexToName[row - 1]
                if itemName not in itemResults:
                    itemResults[itemName] = {}
                itemResults[itemName].update(dataRow)
        # print(itemResults)

    # output a new workbook with each sheet being a item name at A0 and then the data
    newBook = xw.Book()
    for itemName, itemData in itemResults.items():
        sheet = newBook.sheets.add(str(nameToId[itemName]))
        sheet.range("a1").value = itemName
        sheet.range("a2").value = "Date"
        sheet.range("b2").value = "Full Order"
        sheet.range("c2").value = "Short"
        for ind, (date, orders) in enumerate(
            sorted(itemData.items(), key=lambda x: datetime.strptime(x[0], "%m/%d/%Y"))
        ):
            sheet.range(f"a{ind+3}").value = date
            sheet.range(f"b{ind+3}").value = orders[0]
            sheet.range(f"c{ind+3}").value = orders[1]
    newBook.save("./data/output.xlsx")
    print("done")


if __name__ == "__main__":
    main()
