using System.Net.Sockets;

namespace HSCTSV.Model
{
    public class InForStudent
    {
        public string IDStudent { get; set; }
        public string NameStudent { get; set; }
        public string Sexy { get; set; }
        public string BirthDay { get; set; }
        public string Class { get; set; }
        public string Program { get; set; }
        public string Faculty { get; set; }
        public string Email { get; set; }
        public string Semester { get; set; }
        public string TC_All { get; set; }
        public string PointTB { get; set; }
        public string LevelStudent { get; set; }
        public string LevelWarning { get; set; }
        public string Nation { get; set; }
        public string Address { get; set; }
        public string StatusInfor { get; set; }
        public string CMT { get; set; }
        public string Father { get; set; }
        public string Mother { get; set; }
        public string BHYT { get; set; }
    }
    public class HSGetInforByIdStudent
    {
        public int RespCode { get; set; }
        public string RespText { get; set; }
        public InForStudent Infor { get; set; }
    }
}
