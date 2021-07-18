# SMSR_Scrapping

This console application generates the Excel report about the dry land happened in Poland [gminas](https://en.wikipedia.org/wiki/Administrative_divisions_of_Poland) in 2021.

It uses the public [SMSR page](https://susza.iung.pulawy.pl/wykazy/2021/) on which You could find the public data.

Because the page uses server-side rendering I used the HTML-scrapping mechanism to retrieve the data and generate the report.

The report itself generates only columns for categories/crops species for which dry-land happened at least in one gmina of all provinces (voivodeships)
