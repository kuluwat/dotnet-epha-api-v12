using Class;
using dotnet_epha_api.Class;
using Microsoft.AspNetCore.Mvc;
using Model;
 

namespace Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class MasterDataController : ControllerBase
    {
        //private readonly IAuthenticationService _authenticationService;
        //public MasterDataController(IAuthenticationService authenticationService)
        //{
        //    _authenticationService = authenticationService;
        //}


        //[ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        //[HttpPost("get_databasecollation;", Name = "get_databasecollation")]
        //public string get_databasecollation()
        //{
        //    string token_log = "";
        //    string msg = "";
        //    try
        //    {
        //        ClassMasterData cls = new ClassMasterData();
        //        return cls.get_databasecollation();
        //    }
        //    catch (Exception e)
        //    {
        //        msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
        //    }
        //    return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        //}

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("get_manageuser", Name = "get_manageuser")]
        public string get_manageuser(LoadMasterPageModel param)
        {
            string token_log = "";
            string msg = "";
            try
            {

                ClassMasterData cls = new ClassMasterData();
                return cls.get_manageuser(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("set_manageuser", Name = "set_manageuser")]
        public string set_manageuser(SetManageuser param)
        {
            string token_log = "";
            string msg = "";
            try
            {

                ClassMasterData cls = new ClassMasterData();
                return cls.set_manageuser(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("get_authorizationsetting", Name = "get_authorizationsetting")]
        public string get_authorizationsetting(LoadMasterPageModel param)
        {
            string token_log = "";
            string msg = "";
            try
            {

                ClassMasterData cls = new ClassMasterData();
                return cls.get_authorizationsetting(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("set_authorizationsetting", Name = "set_authorizationsetting")]
        public string set_authorizationsetting(SetAuthorizationSetting param)
        {
            string token_log = "";
            string msg = "";
            try
            {

                ClassMasterData cls = new ClassMasterData();
                return cls.set_authorizationsetting(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        #region master systemwide

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("get_master_company", Name = "get_master_company")]
        public string get_master_company(LoadMasterPageModel param)
        {
            string token_log = "";
            string msg = "";
            try
            {
                ClassMasterData cls = new ClassMasterData();
                return cls.get_master_company(param);

            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        //area 
        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("get_master_area", Name = "get_master_area")]
        public string get_master_area(LoadMasterPageModel param)
        {
            string token_log = "";
            string msg = "";
            try
            {

                ClassMasterData cls = new ClassMasterData();
                return cls.get_master_area(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("set_master_area", Name = "set_master_area")]
        public string set_master_area(SetDataMasterModel param)
        {
            string token_log = "";
            string msg = "";
            try
            {

                param.page_name = "area";
                return _set_master_data(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        //toc 
        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("get_master_toc", Name = "get_master_toc")]
        public string get_master_toc(LoadMasterPageModel param)
        {
            string token_log = "";
            string msg = "";
            try
            {

                ClassMasterData cls = new ClassMasterData();
                return cls.get_master_toc(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("set_master_toc", Name = "set_master_toc")]
        public string set_master_toc(SetDataMasterModel param)
        {
            string token_log = "";
            string msg = "";
            try
            {
                param.page_name = "toc";
                return _set_master_data(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        //unit 
        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("get_master_unit", Name = "get_master_unit")]
        public string get_master_unit(LoadMasterPageModel param)
        {
            string token_log = "";
            string msg = "";
            try
            {

                ClassMasterData cls = new ClassMasterData();
                return cls.get_master_unit(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("set_master_unit", Name = "set_master_unit")]
        public string set_master_unit(SetDataMasterModel param)
        {
            string token_log = "";
            string msg = "";
            try
            {

                param.page_name = "unit";
                return _set_master_data(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }
        private string _set_master_data(SetDataMasterModel param)
        {
            string msg = "";
            try
            {
                ClassMasterData cls = new ClassMasterData();
                return cls.set_master_systemwide(param);
            }
            catch (Exception e) { msg = e.Message.ToString(); }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }
        #endregion master systemwide


        #region jsea 
        //task type

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("get_master_tasktype", Name = "get_master_tasktype")]
        public string get_master_tasktype(LoadMasterPageModel param)
        {
            string token_log = "";
            string msg = "";
            try
            {

                ClassMasterData cls = new ClassMasterData();
                return cls.get_master_tasktype(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("set_master_tasktype", Name = "set_master_tasktype")]
        public string set_master_tasktype(SetDataMasterModel param)
        {
            string token_log = "";
            string msg = "";
            try
            {
                ClassMasterData cls = new ClassMasterData();
                return cls.set_master_tasktype(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }


        //Tag ID/Equipment
        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("get_master_tagid", Name = "get_master_tagid")]
        public string get_master_tagid(LoadMasterPageModel param)
        {
            string token_log = "";
            string msg = "";
            try
            {
                ClassMasterData cls = new ClassMasterData();
                return cls.get_master_tagid(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("set_master_tagid", Name = "set_master_tagid")]
        public string set_master_tagid(SetDataMasterModel param)
        {
            string token_log = "";
            string msg = "";
            try
            {
                ClassMasterData cls = new ClassMasterData();
                return cls.set_master_tagid(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        //mandatory_note
        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("get_master_mandatorynote", Name = "get_master_mandatorynote")]
        public string get_master_mandatorynote(LoadMasterPageModel param)
        {
            string token_log = "";
            string msg = "";
            try
            {
                ClassMasterData cls = new ClassMasterData();
                return cls.get_master_mandatorynote(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("set_master_mandatorynote", Name = "set_master_mandatorynote")]
        public string set_master_mandatorynote(SetDataMasterModel param)
        {
            string token_log = "";
            string msg = "";
            try
            {
                ClassMasterData cls = new ClassMasterData();
                return cls.set_master_mandatorynote(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }
        #endregion jsea 

        #region hazop module 
        //Functional Location

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("get_master_functionallocation", Name = "get_master_functionallocation")]
        public string get_master_functionallocation(LoadMasterPageModel param)
        {
            string token_log = "";
            string msg = "";
            try
            {
                ClassMasterData cls = new ClassMasterData();
                return cls.get_master_functionallocation(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("set_master_functionallocation", Name = "set_master_functionallocation")]
        public string set_master_functionallocation(SetMasterGuideWordsModel param)
        {
            string token_log = "";
            string msg = "";
            try
            {
                ClassMasterData cls = new ClassMasterData();
                return cls.set_master_functionallocation(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        //guidewords

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("get_master_guidewords", Name = "get_master_guidewords")]
        public string get_master_guidewords(LoadMasterPageModel param)
        {
            string token_log = "";
            string msg = "";
            try
            {
                ClassMasterData cls = new ClassMasterData();
                return cls.get_master_guidewords(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("set_master_guidewords", Name = "set_master_guidewords")]
        public string set_master_guidewords(SetMasterGuideWordsModel param)
        {
            string token_log = "";
            string msg = "";
            try
            {
                ClassMasterData cls = new ClassMasterData();
                return cls.set_master_guidewords(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }
        #endregion hazop module 

        #region Manage User

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("get_master_contractlist", Name = "get_master_contractlist")]
        public string get_master_contractlist(LoadMasterPageModel param)
        {
            string token_log = "";
            string msg = "";
            try
            {
                ClassMasterData cls = new ClassMasterData();
                return cls.get_master_contractlist(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));

        }

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("set_master_contractlist", Name = "set_master_contractlist")]
        public string set_master_contractlist(SetMasterGuideWordsModel param)
        {
            string token_log = "";
            string msg = "";
            try
            {
                ClassMasterData cls = new ClassMasterData();
                return cls.set_master_contractlist(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        #endregion Manage User


        #region hra
        //get_master_sections_group

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("get_master_sections_group", Name = "get_master_sections_group")]
        public string get_master_sections_group(LoadMasterPageModel param)
        {
            string token_log = "";
            string msg = "";
            try
            {
                ClassMasterData cls = new ClassMasterData();
                return cls.get_master_sections_group(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("set_master_sections_group", Name = "set_master_sections_group")]
        public string set_master_sections_group(SetMasterGuideWordsModel param)
        {
            string token_log = "";
            string msg = "";
            try
            {
                ClassMasterData cls = new ClassMasterData();
                return cls.set_master_sections_group(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        //get_master_sub_area_group
        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("get_master_sub_area_group", Name = "get_master_sub_area_group")]
        public string get_master_sub_area_group(LoadMasterPageModel param)
        {
            string token_log = "";
            string msg = "";
            try
            {
                ClassMasterData cls = new ClassMasterData();
                return cls.get_master_sub_area_group(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }


        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("set_master_sub_area_group", Name = "set_master_sub_area_group")]
        public string set_master_sub_area_group(SetMasterGuideWordsModel param)
        {
            string token_log = "";
            string msg = "";
            try
            {
                ClassMasterData cls = new ClassMasterData();
                return cls.set_master_sub_area_group(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }


        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("get_master_sub_area_equipmet", Name = "get_master_sub_area_equipmet")]
        public string get_master_sub_area_equipmet(LoadMasterPageModel param)
        {
            string token_log = "";
            string msg = "";
            try
            {
                ClassMasterData cls = new ClassMasterData();
                return cls.get_master_sub_area_equipmet(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }


        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("set_master_sub_area_equipmet", Name = "set_master_sub_area_equipmet")]
        public string set_master_sub_area_equipmet(SetMasterGuideWordsModel param)
        {
            string token_log = "";
            string msg = "";
            try
            {
                ClassMasterData cls = new ClassMasterData();
                return cls.set_master_sub_area_equipmet(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        //get_master_hazard_type
        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("get_master_hazard_type", Name = "get_master_hazard_type")]
        public string get_master_hazard_type(LoadMasterPageModel param)
        {
            string token_log = "";
            string msg = "";
            try
            {
                ClassMasterData cls = new ClassMasterData();
                return cls.get_master_hazard_type(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("set_master_hazard_type", Name = "set_master_hazard_type")]
        public string set_master_hazard_type(SetMasterGuideWordsModel param)
        {
            string token_log = "";
            string msg = "";
            try
            {
                ClassMasterData cls = new ClassMasterData();
                return cls.set_master_hazard_type(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        //get_master_hazard_riskfactors
        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("get_master_hazard_riskfactors", Name = "get_master_hazard_riskfactors")]
        public string get_master_hazard_riskfactors(LoadMasterPageModel param)
        {
            string token_log = "";
            string msg = "";
            try
            {
                ClassMasterData cls = new ClassMasterData();
                return cls.get_master_hazard_riskfactors(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("set_master_hazard_riskfactors", Name = "set_master_hazard_riskfactors")]
        public string set_master_hazard_riskfactors(SetMasterGuideWordsModel param)
        {
            string token_log = "";
            string msg = "";
            try
            {
                ClassMasterData cls = new ClassMasterData();
                return cls.set_master_hazard_riskfactors(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        //get_master_group_list 
        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("get_master_group_list", Name = "get_master_group_list")]
        public string get_master_group_list(LoadMasterPageBySectionModel param)
        {
            string token_log = "";
            string msg = "";
            try
            {
                ClassMasterData cls = new ClassMasterData();
                return cls.get_master_group_list(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("set_master_group_list", Name = "set_master_group_list")]
        public string set_master_group_list(SetMasterGuideWordsModel param)
        {
            string token_log = "";
            string msg = "";
            try
            {
                ClassMasterData cls = new ClassMasterData();
                return cls.set_master_group_list(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        //get_master_worker_group 
        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("get_master_worker_group", Name = "get_master_worker_group")]
        public string get_master_worker_group(LoadMasterPageBySectionModel param)
        {
            string token_log = "";
            string msg = "";
            try
            {
                ClassMasterData cls = new ClassMasterData();
                return cls.get_master_worker_group(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("set_master_worker_group", Name = "set_master_worker_group")]
        public string set_master_worker_group(SetMasterGuideWordsModel param)
        {
            string token_log = "";
            string msg = "";
            try
            {
                ClassMasterData cls = new ClassMasterData();
                return cls.set_master_worker_group(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }


        //get_master_worker_list 
        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("get_master_worker_list", Name = "get_master_worker_list")]
        public string get_master_worker_list(LoadMasterPageModel param)
        {
            string token_log = "";
            string msg = "";
            try
            {
                ClassMasterData cls = new ClassMasterData();
                return cls.get_master_worker_list(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        [ValidateAntiForgeryToken] //[ApiAuthorize(Scope.Internal)]
        [HttpPost("set_master_worker_list", Name = "set_master_worker_list")]
        public string set_master_worker_list(SetMasterGuideWordsModel param)
        {
            string token_log = "";
            string msg = "";
            try
            {
                ClassMasterData cls = new ClassMasterData();
                return cls.set_master_worker_list(param);
            }
            catch (Exception e)
            {
                msg += $" method error: {e.Message.ToString()} -> token log: {token_log}";
            }
            return ClassJSON.SetJSONresultRef(ClassFile.refMsgSave((msg == "" ? "true" : "false"), msg, "", "", "", ""));
        }

        #endregion hra

    }
}
