var PC = null;

function PC_Init() {

 PC = new Portal_Calendar();

 PC.Init();

 PC.Add("pc_1", document.formContent.h_TopicStart);
 PC.Add("pc_2", document.formContent.h_TopicExpires);

 PC.SetFirstDayOfWeek("pc_1", 1);
 PC.SetFirstDayOfWeek("pc_2", 1);

 PC.SetPosition("pc_1", 160, 200);
 PC.SetPosition("pc_2", 160, 200);

 PC.SetImage("pc_1", "NEXT", "<img src=\"" + PC_PATH + "Portal_Image/Admin/c_next.gif\" width=\"14\" height=\"14\" border=\"0\">");
 PC.SetImage("pc_1", "PREV", "<img src=\"" + PC_PATH + "Portal_Image/Admin/c_prev.gif\" width=\"14\" height=\"14\" border=\"0\">");
 PC.SetImage("pc_1", "CLOSE", "<img src=\"" + PC_PATH + "Portal_Image/Admin/c_close.gif\" width=\"14\" height=\"14\" border=\"0\">");
 PC.SetImage("pc_2", "NEXT", "<img src=\"" + PC_PATH + "Portal_Image/Admin/c_next.gif\" width=\"14\" height=\"14\" border=\"0\">");
 PC.SetImage("pc_2", "PREV", "<img src=\"" + PC_PATH + "Portal_Image/Admin/c_prev.gif\" width=\"14\" height=\"14\" border=\"0\">");
 PC.SetImage("pc_2", "CLOSE", "<img src=\"" + PC_PATH + "Portal_Image/Admin/c_close.gif\" width=\"14\" height=\"14\" border=\"0\">");

 PC.SetColor("pc_1", "BORDER", "#8183A2");
 PC.SetColor("pc_2", "BORDER", "#8183A2");

 PC.SetDateFormat("pc_1", "dd/mm/yyyy");
 PC.SetDateFormat("pc_2", "dd/mm/yyyy");

 PC.InitCal("pc_1");
 PC.InitCal("pc_2");

};