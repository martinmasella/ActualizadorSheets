public class Ticker 
{
    public string NombreLargo { get; set; }
    public string NombreMedio { get; set; }
    public string NombreCorto { get; set; }
    public decimal BidSize { get; set; }
    public decimal Bid { get; set; }
    public decimal Last { get; set; }
    public decimal Offer { get; set; }
    public decimal OfferSize { get; set; }
    public string Stamp { get; set; }

    public Ticker(string nombrelargo, string nombremedio, string nombrecorto)
    {
        NombreLargo = nombrelargo;
        NombreMedio = nombremedio;
        NombreCorto = nombrecorto;
        BidSize = Bid = Last = Offer = OfferSize = 0;
    }
}