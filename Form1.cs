using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using Primary;
using Primary.Data;

namespace ActualizadorSheets
{
    public partial class Form1 : Form
    {

        const string sURL = "https://api.veta.xoms.com.ar";
        const string prefijoPrimary = "MERV - XMEV - ";
        const string sufijoCI = " - CI";
        //const string sufijo48 = " - 48hs";
        const string sufijo24 = " - 24hs";

        string LEDEPesos;
        string LEDEDolares;

        string token;
        List<string> nombres;
        List<Ticker> tickers;

        //Acá hardcodee el ID de mi Google Sheet.
        String spreadsheetId2 = "1W2kTK4n10-fKWYRJmoffOnxdaibkL-7I5wduSHVlGvI";

        UserCredential credencial;
        SheetsService service;
        //Idem con el archivo jSon del Secret.
        FileStream stream = new FileStream("client_secret_202947654746-fmqgbul2aegj1gqe66cjh5tb633oadcn.apps.googleusercontent.com.json", FileMode.Open, FileAccess.Read);
        string credPath = "token.json";
        static string[] Scopes = { "https://www.googleapis.com/auth/spreadsheets" };

        String range2 = "Dato!B1";

        public Form1()
        {
            InitializeComponent();
            credencial = GoogleWebAuthorizationBroker.AuthorizeAsync(
            GoogleClientSecrets.FromStream(stream).Secrets,
            Scopes,
            "user",
            CancellationToken.None,
            new FileDataStore(credPath, true)).Result;
            service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credencial,
                ApplicationName = "Arbitraje Ratios"
            });
            CheckForIllegalCrossThreadCalls = false;

            var configuracion = new ConfigurationBuilder()
                .SetBasePath(AppContext.BaseDirectory)
                .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
                .Build();
            LEDEPesos = configuracion.GetSection("MiConfiguracion:LEDE Pesos").Value;
            LEDEDolares = configuracion.GetSection("MiConfiguracion:LEDE Dolares").Value;
            txtUsuario.Text = configuracion.GetSection("MiConfiguracion:UsuarioVETA").Value;
            txtClave.Text = configuracion.GetSection("MiConfiguracion:ClaveVETA").Value;
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            // The A1 notation of the values to clear.
            string range = "Intraday!O2:T2550";  // TODO: Update placeholder value.

            // TODO: Assign values to desired properties of `requestBody`:
            ClearValuesRequest requestBody = new ClearValuesRequest();
            SpreadsheetsResource.ValuesResource.ClearRequest request = service.Spreadsheets.Values.Clear(requestBody, spreadsheetId2, range);

            // To execute asynchronously in an async method, replace `request.Execute()` as shown:
            ClearValuesResponse response = request.Execute();
            // Data.ClearValuesResponse response = await request.ExecuteAsync();

            // TODO: Change code below to process the `response` object:
            Console.WriteLine(JsonConvert.SerializeObject(response));
            ToLog("Cleared");
        }

        private void btnLogin_Click(object sender, EventArgs e)
        {
            Inicio();
        }

        private async Task Inicio()
        {
            var api = new Api(new Uri(sURL));
            await api.Login(txtUsuario.Text, txtClave.Text);
            ToLog("Login VETA Ok");

            var allInstruments = await api.GetAllInstruments();

            var entries = new[] { Entry.Last, Entry.Bids, Entry.Offers };

            FillListaTickers();

            var instrumentos = allInstruments.Where(c => tickers.Any(t => t.NombreLargo == c.Symbol));
            using var socket = api.CreateMarketDataSocket(instrumentos, entries, 1, 1);
            socket.OnData = OnMarketData;
            var socketTask = await socket.Start();
            socketTask.Wait(1000);
            ToLog("Websocket Ok");
            tmrRefresh.Interval = 10000;
            tmrRefresh.Enabled = true;
            tmrRefresh.Start();

            await socketTask;
        }

        private async void OnMarketData(Api api, MarketData marketData)
        {
            try
            {
                //En la versión Alpha era así. Gracias Juan Manuel Álvarez.
                //var ticker = marketData.Instrument.Symbol;
                var ticker = marketData.InstrumentId.Symbol;
                decimal bid = 0;
                if (marketData.Data.Bids != null)
                {
                    bid = marketData.Data.Bids.FirstOrDefault().Price;
                }
                decimal offer = 0;
                if (marketData.Data.Offers != null)
                {
                    offer = marketData.Data.Offers.FirstOrDefault().Price;
                }
                decimal last = 0;
                if (marketData.Data.Last != null)
                {
                    last = marketData.Data.Last.Price;
                }

                var elemento = tickers.Where<Ticker>(t => t.NombreLargo == ticker).FirstOrDefault<Ticker>();
                elemento.bid = bid;
                elemento.offer = offer;
                elemento.last = last;

                ToLog(ticker);
            }
            catch (Exception ex)
            {
                ToLog(ex.Message);
            }
        }

        private void FillListaTickers()
        {
            nombres.Add("GD30");
            nombres.Add("AL30");
            nombres.Add("GD30D");
            nombres.Add("AL30D");
            nombres.Add(LEDEPesos);
            nombres.Add(LEDEDolares);
            nombres.Add("CLSIO");

            foreach (var nombre in nombres)
            {
                tickers.Add(new Ticker(prefijoPrimary + nombre + sufijoCI, nombre + "CI", nombre));
                tickers.Add(new Ticker(prefijoPrimary + nombre + sufijo24, nombre + "24", nombre));
            }
        }

        private void ToLog(string s)
        {
            string texto = $"{DateTime.Now.ToLongTimeString()}: {s}";
            lstLog.Items.Add(texto);
            lstLog.SelectedIndex = lstLog.Items.Count - 1;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            this.Top = 10;
            this.Text = "Actualizador Sheets";
            nombres = new List<string>();
            tickers = new List<Ticker>();
        }

        private void tmrRefresh_Tick(object sender, EventArgs e)
        {
            ToLog("Tick");
            string ahora = DateTime.Now.ToString("HH:mm:ss");

            decimal PGD3024 = Precio("GD3024", "Precio");
            decimal PAL3024 = Precio("AL3024", "Precio");
            decimal ratio = 0;
            if (PGD3024 > 0 && PAL3024 > 0)
            {
                ratio = Math.Round(((PGD3024 / PAL3024) - 1) * 100, 2);
            }
            decimal PGD30V = Precio("GD30CI", "Venta");
            decimal PGD30DC = Precio("GD30DCI", "Compra");
            decimal PAL30DV = Precio("AL30DCI", "Venta");
            decimal PAL30C = Precio("AL30CI", "Compra");
            decimal PAL30V = Precio("AL30CI", "Venta");
            decimal PAL30DC = Precio("AL30DCI", "Compra");
            decimal PGD30DV = Precio("GD30DCI", "Venta");
            decimal PGD30C = Precio("GD30CI", "Compra");
            decimal PGD30V24 = Precio("GD3024", "Venta");
            decimal PGD30C24 = Precio("GD3024", "Compra");
            decimal PAL30V24 = Precio("AL3024", "Venta");
            decimal PAL30C24 = Precio("AL3024", "Compra");
            decimal PLEDEV = Precio(LEDEPesos + "CI", "Venta");
            decimal PLEDEC = Precio(LEDEPesos + "CI", "Compra");
            decimal PLEDEDV = Precio(LEDEDolares + "CI", "Venta");
            decimal PLEDEDC = Precio(LEDEDolares + "CI", "Compra");
            decimal PGD30D = Precio("GD30D24", "Precio");
            decimal PAL30D = Precio("AL30D24", "Precio");
            decimal PAL30D24V = Precio("AL30D24", "Venta");
            decimal PAL30D24C = Precio("AL30D24", "Compra");
            decimal PGD30DV24 = Precio("GD30D24", "Venta");
            decimal PGD30DC24 = Precio("GD30D24", "Compra");
            decimal PLEDEV24 = Precio(LEDEPesos + "24", "Venta");
            decimal PLEDEC24 = Precio(LEDEPesos + "24", "Compra");
            decimal PLEDEDV24 = Precio(LEDEDolares + "24", "Venta");
            decimal PLEDEDC24 = Precio(LEDEDolares + "24", "Compra");
            decimal PCLSIOC = Precio("CLSIOCI", "Compra");
            decimal PCLSIOV24 = Precio("CLSIO24", "Venta");


            decimal PuntasGDAL = 0;
            decimal PuntasALGD = 0;
            if (PGD30C24 > 0 && PAL30V24 > 0 && PGD30V24 > 0 && PAL30C24 > 0)
            {
                PuntasGDAL = Math.Round(((PGD30C24 / PAL30V24) - 1) * 100, 2);
                PuntasALGD = Math.Round(((PGD30V24 / PAL30C24) - 1) * 100, 2);
            }

            if (int.Parse(DateTime.Now.ToString("HHmm")) >= 1103 && int.Parse(DateTime.Now.ToString("HHmm")) < 1702)
            {
                ValueRange valueRange = new ValueRange();
                valueRange.MajorDimension = "COLUMNS";//"ROWS";//COLUMNS

                var oblist = new List<object>() { PGD3024, PAL3024, ratio,
                    PGD30V, PGD30DC, PAL30DV, PAL30C, PAL30V, PAL30DC, PGD30DV, PGD30C,
                    PGD30V24, PGD30C24, PAL30V24, PAL30C24, PLEDEV, PLEDEC, PLEDEDV,
                    PLEDEDC, PGD30D, PAL30D,PAL30D24V,PAL30D24C, PGD30DV24, PGD30DC24,
                    PLEDEV24, PLEDEC24, PLEDEDV24, PLEDEDC24, PCLSIOC, PCLSIOV24};

                valueRange.Values = new List<IList<object>> { oblist };

                SpreadsheetsResource.ValuesResource.UpdateRequest update;
                update = service.Spreadsheets.Values.Update(valueRange, spreadsheetId2, range2);
                update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                UpdateValuesResponse result2 = update.Execute();


                //Intraday

                ValueRange lectura = service.Spreadsheets.Values.Get(spreadsheetId2, "Intraday!A1:A2550").Execute();

                int siguiente = lectura.Values.Count + 1;
                valueRange = new ValueRange();
                valueRange.MajorDimension = "ROWS";
                oblist = new List<object>() { ahora, PGD3024, PAL3024, ratio, PuntasGDAL, PuntasALGD };
                valueRange.Values = new List<IList<object>> { oblist };
                update = service.Spreadsheets.Values.Update(valueRange, spreadsheetId2, "Intraday!O" + siguiente.ToString() + ":T" + siguiente.ToString());
                update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
                try
                {
                    update.Execute();

                }
                catch (Exception ex)
                {
                    ToLog(ex.Message);
                }

            }
        }


        private decimal Precio(string nombreMedio, string cual)
        {
            var ticker = tickers.FirstOrDefault(t => t.NombreMedio == nombreMedio);
            if (ticker == null)
            {
                throw new ArgumentException("No se encontró el ticker con el nombre especificado.");
            }
            switch (cual)
            {
                case "Precio":
                    return ticker.last;
                case "Venta":
                    return ticker.offer;
                case "Compra":
                    return ticker.bid;
                default:
                    throw new ArgumentException("El valor " + cual + " no es válido.");
            }
        }


    }
    public class Ticker
    {
        public string NombreLargo { get; set; }
        public string NombreMedio { get; set; }
        public string NombreCorto { get; set; }
        public decimal bid;
        public decimal last;
        public decimal offer;
        public Ticker(string nombrelargo, string nombremedio, string nombrecorto)
        {
            NombreLargo = nombrelargo;
            NombreMedio = nombremedio;
            NombreCorto = nombrecorto;
            bid = 0;
            last = 0;
            offer = 0;
        }
    }
}