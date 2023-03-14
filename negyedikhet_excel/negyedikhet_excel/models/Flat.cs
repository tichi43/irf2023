using System;
using System.Collections.Generic;

#nullable disable

namespace negyedikhet_excel.models
{
    public partial class Flat
    {
        public int FlatSk { get; set; }
        public string Code { get; set; }
        public string Vendor { get; set; }
        public string Side { get; set; }
        public byte District { get; set; }
        public bool Elevator { get; set; }
        public decimal NumberOfRooms { get; set; }
        public short FloorArea { get; set; }
        public decimal Price { get; set; }
    }
}
