const Alpaca = require('@alpacahq/alpaca-trade-api');
const XLSX = require('xlsx');
const apikey=""
const secret="";

const alpaca = new Alpaca({
    keyId: apikey,
    secretKey: secret,
    paper: true,
});

async function gethourlydata(symbol,startdate,enddate){
    
    const bars = alpaca.getBarsV2(symbol, {
        start: startdate,
        end: enddate,
        timeframe: alpaca.newTimeframe(1, alpaca.timeframeUnit.HOUR),
    });

    const workbook = XLSX.utils.book_new();
    
    // Create data array starting with headers
    const data = [
        ['Date','Open','High','Low','Close','Volume',]
    ];

    // Collect all data rows
    for await (let b of bars) {

       //convert timestamp to date in format MM/dd/yyyy
        const date = new Date(b.Timestamp).toLocaleDateString('en-US', { month: '2-digit', day: '2-digit', year: 'numeric' });
        const time = new Date(b.Timestamp).toLocaleTimeString();
        const formattedDate = `${date} ${time}`; 
        
        data.push([formattedDate, b.OpenPrice, b.HighPrice, b.LowPrice, b.ClosePrice, b.Volume]);
    }

    // Create worksheet from the complete data array
    const worksheet = XLSX.utils.aoa_to_sheet(data);
    
    // Add the worksheet to the workbook
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');

    // Save the workbook to a file
    XLSX.writeFile(workbook, `${symbol}_hourly_data.xlsx`);
}

gethourlydata('AMD','2024-02-15','2024-12-08');

