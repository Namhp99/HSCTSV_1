using System;
using System.Collections.Generic;
using System.Net.Sockets;

namespace HSCTSV.Model
{
    public class Paper
    {
        public string TypePaper { get; set; }
        public string Description { get; set; }
        public string FilePath { get; set; }
    }
    public class PaperLst
    {
        public int RespCode { get; set; }
        public string RespText { get; set; }
        public List<Paper> HSPaperLst { get; set; }

    }
    public class HSStudentPaper
    {
        public string TypePaper { get; set; }
        public int IDQuestion { get; set; }
        public string Answer { get; set; }
        public string Question { get; set; }
    }
    public class HSStudentPaperLst
    {
        public int RespCode { get; set; }
        public string RespText { get; set; }
        public List<HSStudentPaper> HSStudentQuestionPaperLst { get; set; }
    }
    public class HSStudentPaperById
    {
        public string TypePaper { get; set; }
        public string Description { get; set; }
        public string TimeCreate { get; set; }
        public string TimeAccept { get; set; }
        public string UserAccept { get; set; }
        public int Status { get; set; }
        public string RowID { get; set; }
        public string Note { get; set; }
        public string IDStudent { get; set; }
        public string NameStudent { get; set; }
        public string MailAccept { get; set; }

        public DateTime? TimeCreateD
        {
            get { return Utils.Utils.convertTimeAPI(TimeCreate); }
        }
        public DateTime? TimeAcceptD
        {
            get { return Utils.Utils.convertTimeAPI(TimeAccept); }
        }
        
    }
    public class HSPaperStudentByIdLst
    {
        public int RespCode { get; set; }
        public string RespText { get; set; }
        public List<HSStudentPaperById> HSPaperStudentInforLst { get; set; }
    }
    public class HSGetPaperByDate
    {
        public int RespCode { get; set; }
        public string RespText { get; set; }
        public List<HSStudentPaperById> HSPaperStudentInforLst { get; set; }
    }

    public class HSPaperDetail
    {
        public static string saveDateStart;
        public static string saveDateEnd;
        public string TypePaper { get; set; }
        public string Description { get; set; }
        public string TimeCreate { get; set; }
        public string TimeAccept { get; set; }
        public string UserAccept { get; set; }
        public int Status { get; set; }
        public string RowID { get; set; }
        public string Note { get; set; }
        public string IDStudent { get; set; }
        public string DescriptionPaper { get; set; }
        public string IDPaper { get; set; }
        public string FileName { get; set; }
        public List<HSStudentPaper> QuestionLst { get; set; }
        public DateTime? TimeCreateD
        {
            get { return Utils.Utils.convertTimeAPI(TimeCreate); }
        }
        public DateTime? TimeAcceptD
        {
            get { return Utils.Utils.convertTimeAPI(TimeAccept); }
        }
    }
    public class HSGetPaperDetailByID
    {
        public int RespCode { get; set; }
        public string RespText { get; set; }
        public HSPaperDetail Data { get; set; }

    }
}
