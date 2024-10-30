using System.ComponentModel.DataAnnotations;

namespace Model
{
    public class PersonDetailsModel
    {
        public string? user_type { get; set; } = "";
    }

    public class LoadMasterPageModel
    { 
        public string? role_type { get; set; } = "";
        public string? user_name { get; set; } = "";
    }
    public class LoadMasterPageBySectionModel
    {
        public string? role_type { get; set; } = "";
        public string? user_name { get; set; } = "";
        public string? id_department { get; set; } = "";
        public string? id_sections { get; set; } = "";
    }
    public class LoadMasterPageByWorkerModel
    { 
        public string? role_type { get; set; } = "";
        public string? user_name { get; set; } = "";
        public string? id_worker_group { get; set; } = ""; 
    }


    public class AreaModel
    {
        public string? role_type { get; set; } = "";
        public string? user_name { get; set; } = "";
        public int? seq { get; set; } = 0;
        public int? id { get; set; } = 0;

        public string? name { get; set; } = "";
        public string? descriptions { get; set; } = "";

        public int? active_type { get; set; } = 0;

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime create_date { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime update_date { get; set; }

        public string? create_by { get; set; } = "";
        public string? update_by { get; set; } = "";
    }
    public class BusinessUnitModel
    {
        public string? role_type { get; set; } = "";
        public string? user_name { get; set; } = "";
        public int? seq { get; set; } = 0;
        public int? id { get; set; } = 0;
        public int? id_area { get; set; } = 0;

        public string? name { get; set; } = "";
        public string? descriptions { get; set; } = "";

        public int? active_type { get; set; } = 0;

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime create_date { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime update_date { get; set; }

        public string? create_by { get; set; } = "";
        public string? update_by { get; set; } = "";
    }
    public class UnitNoModel
    {
        public string? role_type { get; set; } = "";
        public string? user_name { get; set; } = "";
        public int? seq { get; set; } = 0;
        public int? id { get; set; } = 0;
        public int? id_area { get; set; } = 0;
        public int? id_business_unit { get; set; } = 0;

        public string? name { get; set; } = "";
        public string? descriptions { get; set; } = "";

        public int? active_type { get; set; } = 0;

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime create_date { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime update_date { get; set; }

        public string? create_by { get; set; } = "";
        public string? update_by { get; set; } = "";
    }

    //*** master page 
    public class SetDataMasterModel
    {
        public string? role_type { get; set; } = "";
        public string? user_name { get; set; } = "";
        public string? token_doc { get; set; } = "";
        public string? page_name { get; set; } = "";
        public string? json_data { get; set; } = "";
    }
    public class SetManageuser
    {
        public string? user_name { get; set; } = "";
        public string? role_type { get; set; } = "";
        public string? json_register_account { get; set; } = "";
    }
    public class SetAuthorizationSetting
    {
        public string? user_name { get; set; } = "";
        public string? role_type { get; set; } = "";
        public string? json_role_type { get; set; } = "";
        public string? json_menu_setting { get; set; } = "";
        public string? json_role_setting { get; set; } = "";
    } 
    public class DataMasterListModel
    {
        public string? role_type { get; set; } = "";
        public string? user_name { get; set; } = "";
        public string? json_name { get; set; } = "";
        public string? json_data { get; set; } = "";
    }


    //seq,id,no,no_deviations,no_guide_words,deviations,guide_words,process_deviation,area_application,parameter,active_type,def_selected
    public class SetMasterGuideWordsModel
    {
        public string? role_type { get; set; } = "";
        public string? user_name { get; set; } = "";
        public string? json_data { get; set; } = "";
        public string? json_drawing { get; set; } = "";
    }
}
