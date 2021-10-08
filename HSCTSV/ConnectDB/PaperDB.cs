using Flurl.Http;
using HSCTSV.Model;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace HSCTSV.ConnectDB
{
    class PaperDB
    {
        IFlurlClient fc;
        public async Task<PaperLst> GetPaperLst(Dictionary<string, string> objReq)
        {
            fc = new FlurlClient().EnableCookies();
            PaperLst objRes = new PaperLst();
            try
            {
                objRes = await (Utils.Utils.urlAPI + "/Paper/GetPaper")
                             .WithClient(fc).PostUrlEncodedAsync(objReq).ReceiveJson<PaperLst>();
                return objRes;
            }
            catch (Exception ex)
            {
                objRes = new PaperLst
                {
                    RespCode = -2,
                    RespText = ex.Message
                };
                return objRes;
            }
        }
        public async Task<HSStudentPaperLst> GetPaperByRowHeader(Dictionary<string, string> objReq)
        {
            fc = new FlurlClient().EnableCookies();
            HSStudentPaperLst objRes = new HSStudentPaperLst();
            try
            {
                objRes = await (Utils.Utils.urlAPI + "/Paper/GetPaperByRowHeader")
                               .WithClient(fc).PostUrlEncodedAsync(objReq).ReceiveJson<HSStudentPaperLst>();
                return objRes;
            }
            catch (Exception ex)
            {
                objRes = new HSStudentPaperLst
                {
                    RespCode = -2,
                    RespText = ex.Message

                };
                return objRes;
            }
        }
        public async Task<HSPaperStudentByIdLst> GetPaperByIdStudentLst(Dictionary<string, string> objReq)
        {
            fc = new FlurlClient().EnableCookies();
            HSPaperStudentByIdLst objRes = new HSPaperStudentByIdLst();
            try
            {
                objRes = await (Utils.Utils.urlAPI + "/Paper/GetPaperByIdStudent")
                           .WithClient(fc).PostUrlEncodedAsync(objReq).ReceiveJson<HSPaperStudentByIdLst>();
                return objRes;
            }
            catch (Exception ex)
            {
                objRes = new HSPaperStudentByIdLst
                {
                    RespCode = -2,
                    RespText = ex.Message
                };
                return objRes;
            }
        }
        public async Task<HSPaperStudentByIdLst> GetPaperAllByIdStudentLst(Dictionary<string, string> objReq)
        {
            fc = new FlurlClient().EnableCookies();
            HSPaperStudentByIdLst objRes = new HSPaperStudentByIdLst();
            try
            {
                objRes = await (Utils.Utils.urlAPI + "/Paper/GetPaperAllByIdStudent")
                           .WithClient(fc).PostUrlEncodedAsync(objReq).ReceiveJson<HSPaperStudentByIdLst>();
                return objRes;
            }
            catch (Exception ex)
            {
                objRes = new HSPaperStudentByIdLst
                {
                    RespCode = -2,
                    RespText = ex.Message
                };
                return objRes;
            }
        }
        public async Task<HSPaperStudentByIdLst> GetPaperByIdStudentLstAccept(Dictionary<string, string> objReq)
        {
            fc = new FlurlClient().EnableCookies();
            HSPaperStudentByIdLst objRes = new HSPaperStudentByIdLst();
            try
            {
                objRes = await (Utils.Utils.urlAPI + "/Paper/GetPaperByIdStudentAccept")
                             .WithClient(fc).PostUrlEncodedAsync(objReq).ReceiveJson<HSPaperStudentByIdLst>();
                return objRes;
            }
            catch (Exception ex)
            {
                objRes = new HSPaperStudentByIdLst
                {
                    RespCode = -2,
                    RespText = ex.Message
                };
                return objRes;
            }
        }
        public async Task<HSPaperStudentByIdLst> GetPaperAllByToken(Dictionary<string, string> objReq)
        {
            fc = new FlurlClient().EnableCookies();
            HSPaperStudentByIdLst objRes = new HSPaperStudentByIdLst();
            try
            {
                objRes = await (Utils.Utils.urlAPI + "/Paper/GetPaperAllByToken")
                            .WithClient(fc).PostUrlEncodedAsync(objReq).ReceiveJson<HSPaperStudentByIdLst>();
                return objRes;
            }
            catch (Exception ex)
            {
                objRes = new HSPaperStudentByIdLst
                {
                    RespCode = -2,
                    RespText = ex.Message

                };
                return objRes;
            }
        }
        public async Task<HSPaperStudentByIdLst> GetPaperAllNoAccept(Dictionary<string, string> objReq)
        {
            fc = new FlurlClient().EnableCookies();
            HSPaperStudentByIdLst objRes = new HSPaperStudentByIdLst();
            try
            {
                objRes = await (Utils.Utils.urlAPI + "/Paper/GetPaperAllNoAccept")
                            .WithClient(fc).PostUrlEncodedAsync(objReq).ReceiveJson<HSPaperStudentByIdLst>();
                return objRes;
            }
            catch (Exception ex)
            {
                objRes = new HSPaperStudentByIdLst
                {
                    RespCode = -2,
                    RespText = ex.Message

                };
                return objRes;
            }
        }
        public async Task<HSPaperStudentByIdLst> GetPaperByDate(Dictionary<string, string> objReq)
        {
            fc = new FlurlClient().EnableCookies();
            HSPaperStudentByIdLst objRes = new HSPaperStudentByIdLst();
            try
            {
                objRes = await (Utils.Utils.urlAPI + "/Paper/GetPaperByDate")
                            .WithClient(fc).PostUrlEncodedAsync(objReq).ReceiveJson<HSPaperStudentByIdLst>();
                return objRes;
            }
            catch (Exception ex)
            {
                objRes = new HSPaperStudentByIdLst
                {
                    RespCode = -2,
                    RespText = ex.Message

                };
                return objRes;
            }
        }
        public async Task<HSAcceptPaperStudent> AcceptPaperStudent(Dictionary<string, string> objReq)
        {
            fc = new FlurlClient().EnableCookies();
            HSAcceptPaperStudent objRes = new HSAcceptPaperStudent();
            try
            {
                objRes = await (Utils.Utils.urlAPI + "/Paper/AcceptPaperStudent")
                            .WithClient(fc).PostUrlEncodedAsync(objReq).ReceiveJson<HSAcceptPaperStudent>();
                return objRes;
            }
            catch (Exception ex)
            {
                objRes = new HSAcceptPaperStudent
                {
                    RespCode = -2,
                    RespText = ex.Message
                };
                return objRes;
            }
        }
        public async Task<HSAcceptPaperStudent> DonePaperStudent(Dictionary<string, string> objReq)
        {
            fc = new FlurlClient().EnableCookies();
            HSAcceptPaperStudent objRes = new HSAcceptPaperStudent();
            try
            {
                objRes = await (Utils.Utils.urlAPI + "/Paper/DonePaperStudent")
                            .WithClient(fc).PostUrlEncodedAsync(objReq).ReceiveJson<HSAcceptPaperStudent>();
                return objRes;
            }
            catch (Exception ex)
            {
                objRes = new HSAcceptPaperStudent
                {
                    RespCode = -2,
                    RespText = ex.Message
                };
                return objRes;
            }
        }
        public async Task<HSAcceptPaperStudent> PausePaperStudent(Dictionary<string, string> objReq)
        {
            fc = new FlurlClient().EnableCookies();
            HSAcceptPaperStudent objRes = new HSAcceptPaperStudent();
            try
            {
                objRes = await (Utils.Utils.urlAPI + "/Paper/PausePaperStudent")
                            .WithClient(fc).PostUrlEncodedAsync(objReq).ReceiveJson<HSAcceptPaperStudent>();
                return objRes;
            }
            catch (Exception ex)
            {
                objRes = new HSAcceptPaperStudent
                {
                    RespCode = -2,
                    RespText = ex.Message
                };
                return objRes;
            }
        }
        public async Task<HSGetPaperDetailByID> GetPaperDetailByID(Dictionary<string, string> objReq)
        {
            fc = new FlurlClient().EnableCookies();
            HSGetPaperDetailByID objRes = new HSGetPaperDetailByID();
            try
            {
                objRes = await (Utils.Utils.urlAPI + "/Paper/GetPaperDetailByID")
                             .WithClient(fc).PostUrlEncodedAsync(objReq).ReceiveJson<HSGetPaperDetailByID>();
                return objRes;
            }
            catch (Exception ex)
            {
                objRes = new HSGetPaperDetailByID
                {
                    RespCode = -2,
                    RespText = ex.Message
                };
                return objRes;
            }
        }
        public async Task<HSAcceptPaperStudent> DelStudentPaper(Dictionary<string, string> objReq)
        {
            fc = new FlurlClient().EnableCookies();
            HSAcceptPaperStudent objRes = new HSAcceptPaperStudent();
            try
            {
                objRes = await (Utils.Utils.urlAPI + "/Paper/DelPaperStudentByAdmin")
                            .WithClient(fc).PostUrlEncodedAsync(objReq).ReceiveJson<HSAcceptPaperStudent>();
                return objRes;
            }
            catch (Exception ex)
            {
                objRes = new HSAcceptPaperStudent
                {
                    RespCode = -2,
                    RespText = ex.Message
                };
                return objRes;
            }


        }
    }
}
