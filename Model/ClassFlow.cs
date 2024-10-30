using System.ComponentModel.DataAnnotations;

namespace Model
{
    public class tokenModel
    {
        public string? user_name { get; set; } = "";
        public string? role_type { get; set; } = "";
    }
    public class uploadFile
    {
        public string? user_name { get; set; } = "";
        public string? role_type { get; set; } = "";
        public IFormFileCollection file_obj { get; set; }
        public string? file_of { get; set; }
        public string? file_name { get; set; }
        public string? file_seq { get; set; }
        public string? file_part { get; set; }
        public string? file_doc { get; set; }
        public string? sub_software { get; set; } = "";
    }
    public class ReportModel
    {
        public string? user_name { get; set; } = "";
        public string? role_type { get; set; } = "";
        public string? seq { get; set; } = "";
        public string? export_type { get; set; } = "";
        public string? sub_software { get; set; } = "";
        public string? ram_type { get; set; } = "";

    }
    public class ReportByWorksheetModel
    {
        public string? user_name { get; set; } = "";
        public string? role_type { get; set; } = "";
        public string? seq { get; set; } = "";
        public string? export_type { get; set; } = "";
        public string? sub_software { get; set; } = "";
        public string? ram_type { get; set; } = "";
        public string? seq_worksheet { get; set; } = "";

    }
    public class CopyFileModel
    {
        //page_start_first,page_start_second,page_end_first,page_end_second
        public string? user_name { get; set; } = "";
        public string? role_type { get; set; } = "";
        public string? file_name { get; set; } = "";
        public string? file_path { get; set; } = "";
        public string? page_start_first { get; set; } = "";
        public string? page_start_second { get; set; } = "";
        public string? page_end_first { get; set; } = "";
        public string? page_end_second { get; set; } = "";
        public string? sub_software { get; set; } = "";

    }
    public class EmailConfigModel
    {
        public string? user_name { get; set; } = "";
        public string? role_type { get; set; } = "";
        public string? user_email { get; set; } = "";

    }
    public class EmployeeModel
    {
        public string? user_name { get; set; } = "";
        public string? role_type { get; set; } = "";
        public string? user_indicator { get; set; } = "";
        public string? user_filter_text { get; set; } = "";
        public string? max_rows { get; set; } = "";

    }
    public class EmployeeListModel
    {
        public List<EmployeeModel> user_name_list = new List<EmployeeModel> { };
    }
    public class LoadDocModel
    {
        public string? user_name { get; set; } = "";
        public string? role_type { get; set; } = "";
        public string? token_doc { get; set; } = "";
        public string? sub_software { get; set; } = "";
        public string? type_doc { get; set; } = "";
        public string? document_module { get; set; } = "";
        public string? pha_status { get; set; } = "";

    }
    public class LoadDocFollowModel
    {
        public string? user_name { get; set; } = "";
        public string? role_type { get; set; } = "";
        public string? token_doc { get; set; } = "";
        public string? sub_software { get; set; } = "";
        public string? type_doc { get; set; } = "";
        public string? pha_no { get; set; } = "";
        public string? responder_user_name { get; set; } = "";
    }
    public class SetDataWorkflowModel
    {
        public string? user_name { get; set; } = ""; 
        public string? role_type { get; set; } = "";
        public string? token_doc { get; set; } = "";
        public string? pha_status { get; set; } = "";
        public string? pha_version { get; set; } = "";
        public string? action_part { get; set; } = "";

        public string? json_header { get; set; } = "";
        public string? json_general { get; set; } = "";
        public string? json_functional_audition { get; set; } = "";
        public string? json_session { get; set; } = "";
        public string? json_memberteam { get; set; } = "";
        public string? json_approver { get; set; } = "";
        public string? json_drawing { get; set; } = "";

        public string? json_managerecom { get; set; } = "";
        public string? json_ram_level { get; set; } = "";
        public string? json_ram_master { get; set; } = "";

        public string? flow_action { get; set; } = "";
        public string? sub_software { get; set; } = "";

        public string? json_flow_action { get; set; } = "";
        public string? json_drawingworksheet { get; set; } = "";

        //hazop
        public string? json_node { get; set; } = "";
        public string? json_nodedrawing { get; set; } = "";
        public string? json_nodeguidwords { get; set; } = "";
        public string? json_nodeworksheet { get; set; } = "";

        //jsea
        public string? json_tasks_worksheet { get; set; } = "";
        //public string? json_relatedpeople { get; set; } = "";
        //public string? json_relatedpeople_outsider { get; set; } = "";

        //whatif
        public string? json_relatedpeople { get; set; } = "";
        public string? json_relatedpeople_outsider { get; set; } = "";
        public string? json_list { get; set; } = "";
        public string? json_listdrawing { get; set; } = "";
        public string? json_listworksheet { get; set; } = "";


        //hra
        public string? json_subareas { get; set; } = "";
        public string? json_hazard { get; set; } = "";
        public string? json_tasks { get; set; } = "";
        public string? json_descriptions { get; set; } = "";
        public string? json_workers { get; set; } = "";
        public string? json_worksheet { get; set; } = "";
        public string? json_recommendations { get; set; } = "";

    }
    public class ManageDocModel
    {
        public string? user_name { get; set; } = "";
        public string? role_type { get; set; } = "";
        public string? sub_software { get; set; } = "";
        public string? pha_no { get; set; } = "";
        public string? pha_seq { get; set; } = "";
        public string? pha_seq_new { get; set; } = "";
        public string? pha_status_comment { get; set; } = "";
    }
    public class SetDocModel
    {
        public string? user_name { get; set; } = "";
        public string? role_type { get; set; } = "";
        public string? sub_software { get; set; } = "";
        public string? pha_seq { get; set; } = "";
        public string? pha_no { get; set; } = "";
        public string? swhere { get; set; } = "";
    }
    public class HeaderModel
    {
        public string? user_name { get; set; } = "";
        public string? role_type { get; set; } = "";
        public int? seq { get; set; } = 0;
        public int? id { get; set; } = 0;

        public int? year { get; set; } = 0;
        public string? pha_no { get; set; } = "";
        public int? pha_version { get; set; } = 0;
        public int? pha_status { get; set; } = 0;
        public string? pha_request_by { get; set; } = "";
        public string? pha_request_user_name { get; set; } = "";
        public string? pha_request_user_displayname { get; set; } = "";
        public string? pha_sub_software { get; set; } = "";

        public int? request_approver { get; set; } = 0;
        public string? approver_user_name { get; set; } = "";
        public string? approver_user_displayname { get; set; } = "";
        public int? approve_action_type { get; set; } = 0;
        public int? approve_status { get; set; } = 0;
        public string? approve_comment { get; set; } = "";
         
        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime create_date { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime update_date { get; set; }

        public string? create_by { get; set; } = "";
        public string? update_by { get; set; } = "";


        public string? action_type { get; set; } = "";

    }
    public class GeneralModel
    {
        public string? user_name { get; set; } = "";
        public string? role_type { get; set; } = "";
        public int? seq { get; set; } = 0;
        public int? id_pha { get; set; } = 0;
        public int? id { get; set; } = 0;
        public int? id_ram { get; set; } = 0;
        public string? ram { get; set; } = "";
        public string? expense_type { get; set; } = "";
        public string? sub_expense_type { get; set; } = ""; 
        public string? approver_user_name { get; set; } = "";
        public string? approver_user_displayname { get; set; } = "";
        public string? approver_user_img { get; set; } = "";
        public string? reference_moc { get; set; } = "";
        public int? id_area { get; set; } = 0;
        public int? id_business_unit { get; set; } = 0;
        public int? id_unit_no { get; set; } = 0;
        public string? functional_location { get; set; } = "";

        //[target_start_date] date NULL,
        //[target_end_date] date NULL, 
        //[actual_start_date] date NULL,
        //[actual_end_date] date NULL,   
        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime target_start_date { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime target_end_date { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime actual_start_date { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime actual_end_date { get; set; }

        public string? descriptions { get; set; } = "";


        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime create_date { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime update_date { get; set; }

        public string? create_by { get; set; } = "";
        public string? update_by { get; set; } = "";

    }
    public class FunctionalAuditionModel
    {
        public string? user_name { get; set; } = "";
        public string? role_type { get; set; } = "";
        public int? seq { get; set; } = 0;
        public int? id_pha { get; set; } = 0;
        public int? id { get; set; } = 0;
        public string? functional_location { get; set; } = "";

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime create_date { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime update_date { get; set; }

        public string? create_by { get; set; } = "";
        public string? update_by { get; set; } = "";
    }
    public class SessionModel
    {
        public string? user_name { get; set; } = "";
        public string? role_type { get; set; } = "";
        public int? seq { get; set; } = 0;
        public int? id_pha { get; set; } = 0;
        public int? id { get; set; } = 0;


        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime meeting_date { get; set; }
        public string? meeting_start_time { get; set; } = "";
        public string? meeting_end_time { get; set; } = "";

        public int? request_approver { get; set; } = 0;
        public string? approver_user_name { get; set; } = "";
        public string? approver_user_displayname { get; set; } = "";
        public int? approve_action_type { get; set; } = 0;
        public int? approve_status { get; set; } = 0;
        public string? approve_comment { get; set; } = "";

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime create_date { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime update_date { get; set; }

        public string? create_by { get; set; } = "";
        public string? update_by { get; set; } = "";

    }
    public class MemberTeamModel
    {
        public int? seq { get; set; } = 0;
        public int? id_pha { get; set; } = 0;
        public int? id { get; set; } = 0;
         
        public string? user_name { get; set; } = "";
        public string? role_type { get; set; } = "";
        public string? user_displayname { get; set; } = "";

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime create_date { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime update_date { get; set; }

        public string? create_by { get; set; } = "";
        public string? update_by { get; set; } = "";

    }
    public class DrawingModel
    {
        public string? user_name { get; set; } = "";
        public string? role_type { get; set; } = "";
        public int? seq { get; set; } = 0;
        public int? id_pha { get; set; } = 0;
        public int? id { get; set; } = 0;

        public string? document_name { get; set; } = "";
        public string? document_no { get; set; } = "";
        public string? document_file_name { get; set; } = "";
        public string? document_file_path { get; set; } = "";
        public string? descriptions { get; set; } = "";

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime create_date { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime update_date { get; set; }

        public string? create_by { get; set; } = "";
        public string? update_by { get; set; } = "";
    }

    public class NodeModel
    {
        public string? user_name { get; set; } = "";
        public string? role_type { get; set; } = "";
        public int? seq { get; set; } = 0;
        public int? id_pha { get; set; } = 0;
        public int? id { get; set; } = 0;

        public string? node { get; set; } = "";
        public string? design_intent { get; set; } = "";
        public string? design_conditions { get; set; } = "";
        public string? operating_conditions { get; set; } = "";
        public string? node_boundary { get; set; } = "";
        public string? descriptions { get; set; } = "";

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime create_date { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime update_date { get; set; }

        public string? create_by { get; set; } = "";
        public string? update_by { get; set; } = "";
    }
    public class NodeDrawingModel
    {
        public string? user_name { get; set; } = "";
        public string? role_type { get; set; } = "";
        public int? seq { get; set; } = 0;
        public int? id_pha { get; set; } = 0;
        public int? id { get; set; } = 0;

        public int? id_node { get; set; } = 0;
        public int? id_drawing { get; set; } = 0;
        public int? page_start_first { get; set; } = 0;
        public int? page_end_first { get; set; } = 0;
        public int? page_start_second { get; set; } = 0;
        public int? page_end_second { get; set; } = 0;
        public int? page_start_third { get; set; } = 0;
        public int? page_end_third { get; set; } = 0;

        public string? descriptions { get; set; } = "";

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime create_date { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime update_date { get; set; }

        public string? create_by { get; set; } = "";
        public string? update_by { get; set; } = "";
    }

    public class NodeGuidWordsModel
    {
        public string? user_name { get; set; } = "";
        public string? role_type { get; set; } = "";
        public int? seq { get; set; } = 0;
        public int? id_pha { get; set; } = 0;
        public int? id { get; set; } = 0;

        public int? id_node { get; set; } = 0;
        public int? id_guide_word { get; set; } = 0;

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime create_date { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime update_date { get; set; }

        public string? create_by { get; set; } = "";
        public string? update_by { get; set; } = "";
    }
    public class NodeWorksheetModel
    {
        public string? user_name { get; set; } = "";
        public string? role_type { get; set; } = "";
        public int? seq { get; set; } = 0;
        public int? id_pha { get; set; } = 0;
        public int? id { get; set; } = 0;

        public int? id_node { get; set; } = 0;
        public int? id_guide_word { get; set; } = 0;
         
        public int? causes_no { get; set; } = 0;
        public string? causes { get; set; } = "";
        public int? consequences_no { get; set; } = 0;
        public string? consequences { get; set; } = "";
        public int? category_no { get; set; } = 0;
        public string? category_type { get; set; } = "";


        public string? security_befor { get; set; } = "";
        public string? likelihood_befor { get; set; } = "";
        public string? risk_befor { get; set; } = "";
        public string? major_accident_event { get; set; } = "";
        public string? existing_safeguards { get; set; } = "";
        public string? security_after { get; set; } = "";
        public string? likelihood_after { get; set; } = "";
        public string? risk_after { get; set; } = "";
        public string? recommendations { get; set; } = "";
        public string? responder_user_name { get; set; } = "";
        public string? responder_user_displayname { get; set; } = "";
        public string? action_status { get; set; } = "";


        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime create_date { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime update_date { get; set; }

        public string? create_by { get; set; } = "";
        public string? update_by { get; set; } = "";
    }

    public class ManageRecomModel
    {
        public string? user_name { get; set; } = "";
        public string? role_type { get; set; } = "";
        public int? seq { get; set; } = 0;
        public int? id_pha { get; set; } = 0;
        public int? id { get; set; } = 0;

        public string? responder_user_name { get; set; } = "";
        public string? responder_user_displayname { get; set; } = "";
        public string? risk_befor { get; set; } = "";
        public string? risk_after { get; set; } = "";

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime meeting_date { get; set; }
        public string? estimated_start_time { get; set; } = "";
        public string? estimated_end_time { get; set; } = "";
        public string? document_file_name { get; set; } = "";
        public string? document_file_path { get; set; } = "";
        public string? action_status { get; set; } = "";
        public int? responder_action_type { get; set; } = 0;

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime create_date { get; set; }

        [DataType(DataType.Date)]
        [DisplayFormat(DataFormatString = "{0:MM/dd/yyyy}")]
        public DateTime update_date { get; set; }

        public string? create_by { get; set; } = "";
        public string? update_by { get; set; } = "";
    }


    public class StorageLocationModel
    {
        public string? user_name { get; set; } = "";
        public string? role_type { get; set; } = "";
        public int? selected_type { get; set; } = 0;

        public int? id_company { get; set; } = 0;
        public string? name_company { get; set; } = "";
        public int? id_area { get; set; } = 0;
        public string? name_area { get; set; } = "";
        public int? id_apu { get; set; } = 0;
        public string? name_apu { get; set; } = "";
        public int? id_toc { get; set; } = 0;
        public string? name_toc { get; set; } = "";
        public int? id_business_unit { get; set; } = 0;
        public string? name_business_unit { get; set; } = "";
        public int? id_unit_no { get; set; } = 0;
        public string? name_unit_no { get; set; } = "";

    }
    public class GuideWordsModel
    {
        public string? user_name { get; set; } = "";
        public string? role_type { get; set; } = "";
        //deviations, guide_words, process_deviation, area_applocation, 0 as selected_type
        public int? selected_type { get; set; } = 0;
        public string? parameter { get; set; } = "";
        public string? deviations { get; set; } = "";
        public string? guide_words { get; set; } = "";
        public string? process_deviation { get; set; } = "";
        public string? area_application { get; set; } = "";

    }
    public class SuggestionCausesModel
    {
        public string? user_name { get; set; } = "";
        public string? role_type { get; set; } = "";
        public int? selected_type { get; set; } = 0;

        public string? deviations { get; set; } = "";
        public string? guide_words { get; set; } = "";

        public string? causes { get; set; } = "";

    }
    public class SuggestionRecommendationsModel
    {
        public string? user_name { get; set; } = "";
        public string? role_type { get; set; } = "";
        public int? selected_type { get; set; } = 0;

        public string? deviations { get; set; } = "";
        public string? guide_words { get; set; } = "";

        public string? causes { get; set; } = "";
        public string? recommendations { get; set; } = "";

    }
    public class SetDocWorksheetModel
    {
        public string? user_name { get; set; } = "";
        public string? role_type { get; set; } = ""; 
        public string? token_doc { get; set; } = "";
        public string? action { get; set; } = "";
        public string? pha_status { get; set; } = "";
        public string? sub_software { get; set; } = "";
         
        public string? json_worksheet { get; set; } = "";

    }
    public class SetDocApproveModel
    {
        public string? user_name { get; set; } = "";
        public string? role_type { get; set; } = ""; 
        public string? token_doc { get; set; } = "";
        public string? action { get; set; } = "";
        public string? pha_status { get; set; } = "";
        public string? sub_software { get; set; } = "";

        public int? id_session { get; set; } = 0;
        public int? seq { get; set; } = 0;
        public string? action_review { get; set; } = "";
        public string? action_status { get; set; } = "";
        public string? comment { get; set; } = "";
        public string? user_approver { get; set; } = "";

        public string? json_approver { get; set; } = "";
        public string? json_drawing_approver { get; set; } = ""; 

    }
    public class SetDocApproveTa3Model
    {
        public string? user_name { get; set; } = "";
        public string? role_type { get; set; } = "";
        public string? token_doc { get; set; } = "";
        public string? action { get; set; } = "";
        public string? sub_software { get; set; } = "";

        public string? json_header { get; set; } = "";
        public string? json_approver { get; set; } = "";
        public string? json_approver_ta3 { get; set; } = "";

    }
    //set_transfer_monitoring, header, recom_setting, recom_follow
    public class SetDocTransferMonitoringModel
    {
        public string? role_type { get; set; } = "";
        public string? user_name { get; set; } = "";
        public string? token_doc { get; set; } = "";
        public string? action { get; set; } = "";
        public string? sub_software { get; set; } = "";

        public string? json_header { get; set; } = "";
        public string? json_recom_setting { get; set; } = "";
        public string? json_recom_follow { get; set; } = "";

    }
}