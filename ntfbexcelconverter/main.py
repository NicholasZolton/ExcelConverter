import xlwings as xw


def main():
    with xw.Book("./data/data.xlsx", mode="r") as book:
        # figure out the item ids
        firstWeek = book.sheets[4]
        print(firstWeek.name)

        # make a map of index to item id
        def getItemId(index):
            return firstWeek.range("b4:c12").value[index][0]

        sheets = list(book.sheets)[4:]
        for sheet in sheets:
            print(sheet.name)

            # get D3:O12
            data = sheet.range("d3:o12").value
            print(data)

            # map the first row (dates) to "m/d/y"
            dates = data[0]
            for i, date in enumerate(dates):
                dates[i] = date.strftime("%m/%d/%Y")
            print(dates)

            # if col % 2 == 1: then the row is a data row
            for row in range(0, len(data)):
                print(data[row])


if __name__ == "__main__":
    main()
