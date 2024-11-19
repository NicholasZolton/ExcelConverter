import xlwings as xw


def main():
    with xw.Book("./data/data.xlsx", mode="r") as book:
        # figure out the item ids
        firstWeek = book.sheets[4]
        # print(firstWeek.name)

        # parse item ids
        itemIds = list(firstWeek.range("b4:c12").value)
        idToName = {}
        indexToId = {}
        indexToName = {}
        for i, (itemId, itemName) in enumerate(itemIds):
            idToName[int(itemId)] = itemName
            indexToId[i] = itemId
            indexToName[i] = itemName

        itemResults = {}
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
                itemResults[itemName] = itemResults.get(itemName, []) + [dataRow]
        print(itemResults)


if __name__ == "__main__":
    main()
