using System;
namespace GecisSistemi
{
    class Kisi
    {
        public string Ad;
        public string Soyad;
        public string KartId;
        public DateTime GirisDateTime;
        public DateTime CikisDateTime;
        public string GirisTarihi;
        public string CikisTarihi;
        public string SureFarkDakika;
        public bool GirdiMi;
        public string Durum;
        public Kisi(string ad,string soyad,string kartId)
        {
            Ad = ad;
            Soyad = soyad;
            KartId = kartId;
            GirdiMi = false;
        }
    }
}
