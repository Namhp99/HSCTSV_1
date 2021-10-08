using System.Collections.Generic;

namespace HSCTSV.Model
{
    public class HSStudentInfo
    {
        public string UNumberId { get; set; }
        public string UFullName { get; set; }
        public string Usex { get; set; }
        public string UAddress { get; set; }

        public string UCity { get; set; }
        public string UClass { get; set; }
        public string UFaculty { get; set; }
        public string UGroupLevel { get; set; }
        public string UNation { get; set; }
        public string UBirthDay { get; set; }
        public string YearStart { get; set; }
        public string Program { get; set; }
        public string LevelProgram { get; set; }
        public string Email { get; set; }
    }

    public class HSStudentCheckInDevice
    {
        public string UNumberId { get; set; }
        public string UserName { get; set; }
        public string Usex { get; set; }
        public string UAddress { get; set; }
        public string UClassCode { get; set; }
        public string UClass { get; set; }
        public string UMajors { get; set; }
        public string UCourse { get; set; }
        public string UBirthday { get; set; }
        public string UCardId1 { get; set; }
        public string UCardId2 { get; set; }
    }
    public class HSGetUCourseLst
    {
        public int RespCode { get; set; }
        public string RespText { get; set; }
        public List<string> HSUCourseLst { get; set; }
    }
    public class HSGetUMajorsLst
    {
        public int RespCode { get; set; }
        public string RespText { get; set; }
        public List<string> HSUMajorsLst { get; set; }
    }
    public class HSGetStudentLst
    {
        public int RespCode { get; set; }
        public string RespText { get; set; }
        public HSStudentInfo HSStudentInfo { get; set; }

    }

    public class HSGetStudentCheckInDeviceLst
    {
        public int RespCode { get; set; }
        public string RespText { get; set; }
        public List<HSStudentCheckInDevice> HSStudentCheckInDeviceLst { get; set; }
    }
}
