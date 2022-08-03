# XlsxToDataSet-Public

# Usage (C# console application)

// instantiate the handler and load the file

var xl = new Seamlex.Utilities.XlsxToDataSet();

var ds = xl.GetDataSetFromXlsx(@"c:\temp\xlsxtodataset-source.xlsx");


// That's it.  You now have a System.DataSet object with tables



// If your XLSX contains one sheet with the following data:

// id|code|name

// 1|ALPHA|Abc

// 2|BRAVO|def

// 3|CHARLIE|ghi

// 4|DELTA|jkl


// You can test your data as follows:
Console.WriteLine($"ds contains {ds.Tables.Count} tables.");

Console.WriteLine($"ds[0].dt contains {ds.Tables[0].Rows.Count} rows.");

foreach(System.Data.DataRow row in ds.Tables[0].Rows)

    Console.WriteLine($"ds[0].dt.Row col0 = {row.ItemArray[0]}, col1 = {row.ItemArray[1]}, col2 = {row.ItemArray[2]}.");
    


// console output:

ds contains 1 tables.

ds[0].dt contains 4 rows.

ds[0].dt.Row col0 = 1, col1 = ALPHA, col2 = Abc.

ds[0].dt.Row col0 = 2, col1 = BRAVO, col2 = def.

ds[0].dt.Row col0 = 3, col1 = CHARLIE, col2 = ghi.

ds[0].dt.Row col0 = 4, col1 = DELTA, col2 = jkl.

