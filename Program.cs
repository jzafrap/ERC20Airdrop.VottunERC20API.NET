using ExcelLibrary.SpreadSheet;
using VottumERC20API.NET;




//set your api and authorization key
var apiKey = "your-api-key";
var applicationID = "your-application-ID";


//instance a client by calling the provided factory
var vottunAPIForNETClient = VottunApiClientFactory.Create(apiKey, applicationID);



//1. Deploy new ERC20 token
var deployRequest = new DeployRequest
{
    name = "Vottun airdropped Token",
    symbol = "VTNFREE1",
    alias = "Vottun airdropped token ERC20 test1",
    initialSupply = 21000000,
    network = 80002,

};

var deployResponse = await vottunAPIForNETClient.DeployAsync(deployRequest, CancellationToken.None);
Console.WriteLine($"Deployed ContractAddress: {deployResponse.contractAddress}");
var ERC20 = deployResponse.contractAddress;

//2. Open excel file and iterate every row
string ruta = @"C:\Temp\addresslist.xls";
Workbook book = Workbook.Load(ruta);
var sheet = book.Worksheets[0];

var count = sheet.Cells.LastRowIndex;


System.Console.WriteLine("Number of addresses to airdrop:" + count);


for (int rowIndex = sheet.Cells.FirstRowIndex;
       rowIndex <= sheet.Cells.LastRowIndex;
       rowIndex++)
{
    var row = sheet.Cells.GetRow(rowIndex);
    //address to airdrop
    var address = row.GetCell(0).StringValue;

    //Prepare Transfer request via Vottum API for .NET client
    var transferRequest = new TransferRequest
    {
        contractAddress = ERC20,
        recipient = address,
        amount = 100,
        network = 80002,
    };

    try
    {
        var transferResponse = await vottunAPIForNETClient.TransferAsync(transferRequest, CancellationToken.None);
        Console.WriteLine($"Transfer tokens to {transferRequest.recipient} ,TxnHash: {transferResponse.txHash}");
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Txn error:{ex.Message}");
    }

}

Console.WriteLine($"End of airdrop!");