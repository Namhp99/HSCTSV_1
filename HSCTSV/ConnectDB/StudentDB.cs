using Flurl.Http;
using HSCTSV.Model;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace HSCTSV.ConnectDB
{
    class StudentDB
    {
        IFlurlClient fc;
        public async Task<HSGetUMajorsLst> GetUMajorsLst(Dictionary<string, string> objReq)
        {
            fc = new FlurlClient().EnableCookies();
            HSGetUMajorsLst objRes = new HSGetUMajorsLst();
            try
            {
                objRes = await (Utils.Utils.urlAPI + "/Student/GetUMajorsLst")
                                 .WithClient(fc).PostUrlEncodedAsync(objReq).ReceiveJson<HSGetUMajorsLst>();
                return objRes;
            }
            catch (Exception ex)
            {
                objRes = new HSGetUMajorsLst
                {
                    RespCode = -2,
                    RespText = ex.Message
                };
                return objRes;
            }
        }
        public async Task<HSGetUCourseLst> GetUCouseLst(Dictionary<string, string> objReq)
        {
            fc = new FlurlClient().EnableCookies();
            HSGetUCourseLst objRes = new HSGetUCourseLst();
            try
            {
                objRes = await (Utils.Utils.urlAPI + "/Student/GetUCourseLst")
                            .WithClient(fc).PostUrlEncodedAsync(objReq).ReceiveJson<HSGetUCourseLst>();
                return objRes;
            }
            catch (Exception ex)
            {
                objRes = new HSGetUCourseLst
                {
                    RespCode = -2,
                    RespText = ex.Message

                };
                return objRes;
            }
        }
        public async Task<HSGetStudentLst> GetStudentLst(Dictionary<string, string> objReq)
        {
            fc = new FlurlClient().EnableCookies();
            HSGetStudentLst objRes = new HSGetStudentLst();
            try
            {
                objRes = await (Utils.Utils.urlAPI + "/Student/GetStudentLst")
                             .WithClient(fc).PostUrlEncodedAsync(objReq).ReceiveJson<HSGetStudentLst>();
                return objRes;

            }
            catch (Exception ex)
            {
                objRes = new HSGetStudentLst
                {
                    RespCode = -2,
                    RespText = ex.Message
                };
                return objRes;
            }
        }
        public async Task<HSGetStudentLst> GetStudentById(Dictionary<string, string> objReq)
        {
            fc = new FlurlClient().EnableCookies();
            HSGetStudentLst objRes = new HSGetStudentLst();
            try
            {
                objRes = await (Utils.Utils.urlAPI + "/Student/GetStudentById")
                             .WithClient(fc).PostUrlEncodedAsync(objReq).ReceiveJson<HSGetStudentLst>();
                return objRes;

            }
            catch (Exception ex)
            {
                objRes = new HSGetStudentLst
                {
                    RespCode = -2,
                    RespText = ex.Message
                };
                return objRes;
            }
        }
        public async Task<HSGetStudentCheckInDeviceLst> GetStudentCheckInDeviceLst(Dictionary<string, string> objReq)
        {
            fc = new FlurlClient().EnableCookies();
            HSGetStudentCheckInDeviceLst objRes = new HSGetStudentCheckInDeviceLst();
            try
            {
                objRes = await (Utils.Utils.urlAPI + "/Student/GetStudentCheckInDeviceLst")
                             .WithClient(fc).PostUrlEncodedAsync(objReq).ReceiveJson<HSGetStudentCheckInDeviceLst>();
                return objRes;
            }
            catch (Exception ex)
            {
                objRes = new HSGetStudentCheckInDeviceLst
                {
                    RespCode = -2,
                    RespText = ex.Message
                };
                return objRes;
            }
        }
        public async Task<HSGetInforByIdStudent> GetInfoByIdStudent(Dictionary<string, string> objReq)
        {
            fc = new FlurlClient().EnableCookies();
            HSGetInforByIdStudent objRes = new HSGetInforByIdStudent();
            try
            {
                objRes = await (Utils.Utils.urlAPI + "/Student/GetInForStudentById")
                             .WithClient(fc).PostUrlEncodedAsync(objReq).ReceiveJson<HSGetInforByIdStudent>();
                return objRes;
            }
            catch (Exception ex)
            {
                objRes = new HSGetInforByIdStudent
                {
                    RespCode = -2,
                    RespText = ex.Message
                };
                return objRes;
            }

        }
    }
}
