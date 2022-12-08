using EppPlus_Speed_Test.Source;

Console.WriteLine("" +
    "The program is a test of using Epplus via LoadFromCollection. \n" +
    "The excel file will be saved to the desktop into the folder epplus \n");

Console.WriteLine("" +
    "Enter " +
    "the maximum number of rows 1_048_576 and \n" +
    "maximum number of columns 16_384");

Console.WriteLine("" +
    "Enter for example \n" +
    "100_000 rows and \n" +
    "10 columns \n" +
    "in total, 1_000_000 values will be filled in\n");


do
{
    int rowCount;
    int columnCount;

    do
    {
        Console.Write("Enter row count: ");

        var rowInput = Console.ReadLine();

        var isParsed = int.TryParse(rowInput, out rowCount);

        if (isParsed) break;

        Console.WriteLine("Row count is invalid, try again");

    } while (true);

    do
    {
        Console.Write("Enter column number: ");

        var columnInput = Console.ReadLine();

        var isParsed = int.TryParse(columnInput, out columnCount);

        if (isParsed) break;

        Console.WriteLine("Column count is invalid, try again");

    } while (true);


    var array = Excel.CreateArray(rowCount, columnCount);

    Excel.StartClock();

    Excel.Save(array);

    Excel.StopClock();

    Console.WriteLine($"You spent {Excel.Elapsed()} for loading {rowCount * columnCount} values.\n");


    do
    {
        Console.Write("Try again? (y/n): ");

        var exitInput = Console.ReadLine();

        if (exitInput != "y" && exitInput != "n")
        {
            Console.WriteLine("Wrong letter, try again\n");
        }

        if (exitInput == "n") Environment.Exit(0);
        if (exitInput == "y") break;

    } while (true);

} while (true);