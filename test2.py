monthsPossible = ["December2018","January2019", "February2019", "March2019", "April2019", "May2019", "June2019", "July2019", "August2019", "September2019", "October2019", "November2019", "December2019", "January2020", "February2020", "March2020", "April2020", "May2020", "June2020", "July2020", "August2020", "September2020", "October2020", "November2020", "December2020", "January2021", "February2021", "March2021", "April2021", "May2021", "June2021", "July2021", "August2021", "September2021", "October2021"]
monthsNumberToWord = {
        1: "January",
        2: "February",
        3: "March",
        4: "April",
        5: "May",
        6: "June",
        7: "July",
        8: "August",
        9: "September",
        10: "October",
        11: "November",
        12: "December"
}
years = [2019,2020,2021]
string = "["
for year in years:
    for month in monthsNumberToWord.values():
        string = string + '"' + (str(month) + str(year)) + '", '
print(string)