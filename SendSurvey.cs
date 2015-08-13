using System;
using System.ComponentModel;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using Microsoft.SharePoint;
using Microsoft.SharePoint.WebControls;
using Microsoft.SharePoint.Utilities;
using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Text.RegularExpressions;
using System.Collections.Specialized;
using System.Net.Mail;

namespace SendSurvey
{
    [ToolboxItemAttribute(false)]
    public class SendSurvey : WebPart
    {
        DataGrid tb1, tb2;
        DataTable dt, dt2;
        CheckBoxList list;
        List<Map> allMaps;
        List<Map> allGroupMaps;
        ControlCollection contr;
        DropDownList _sourceDropDownList;
        TextBox emailSubject = null, emailBody = null, addRecipients;
        List<String> addedEmails;
        List<SPUser> mailingList;
        ListItem selectedSurvey = null;

        public BoundColumn CreateBoundColumn(String dataFieldValue, String headerTextValue)
        {
            BoundColumn column = new BoundColumn();
            column.DataField = dataFieldValue;
            column.HeaderText = headerTextValue;
            return column;
        }
        protected override void CreateChildControls()
        {
            addedEmails = new List<String>();
            mailingList = new List<SPUser>();

            Label outputOne = new Label();
            outputOne.Text = "Select a survey to issue:<br /><br />";
            this.Controls.Add(outputOne);

            //Create a dropdownlist object
            _sourceDropDownList = new DropDownList();
            _sourceDropDownList.ID = "_sourceDropDownList";

            //Sets up aspects of the dropdownlist
            LoadBoxData();
            this.Controls.Add(_sourceDropDownList);

            Label outputFour = new Label();
            outputFour.Text = "<br /><br />Select the users and groups to send the survey to:<br /><br />";
            this.Controls.Add(outputFour);
            tb1 = new DataGrid();
            tb1.AutoGenerateColumns = false;
            list = new CheckBoxList();
            SPUserCollection userList = SPContext.Current.Web.AllUsers;
            SPGroupCollection groupList = SPContext.Current.Web.Groups;

            //Create a new TemplateColumn object
            TemplateColumn tcol = new TemplateColumn();
            {
                tcol.HeaderText = "Select";
                tcol.ItemTemplate = new DynamicItemTemplate_2();
            }
            DynamicItemTemplate_2 copy = (DynamicItemTemplate_2)tcol.ItemTemplate;
            copy.SetControls(this.Controls);
            copy.SetCheckBoxes(list);
            allMaps = new List<Map>();
            copy.SetMap(allMaps);

            tb1.Columns.Add(tcol);
            tb1.Columns.Add(CreateBoundColumn("StringValue", "User"));
            tb1.Columns.Add(CreateBoundColumn("EmailValue", "Email"));

            if (!Page.IsPostBack)
            {
                tb1.DataBind();
            }

            //Set up data
            dt = new DataTable();
            CheckBoxField tmpField = new CheckBoxField();
            dt.Columns.Add(new DataColumn("IntegerValue", tmpField.GetType()));
            dt.Columns.Add(new DataColumn("StringValue", "".GetType()));
            dt.Columns.Add(new DataColumn("EmailValue", "".GetType()));

            DataRow dr;
            Int16 k;
            if (userList.Count > 0)
            {
                for (k = 0; k < userList.Count; k++)
                {
                    dr = dt.NewRow();
                    CheckBoxField field = new CheckBoxField();
                    dr[0] = field;
                    dr[1] = userList[k].Name;
                    dr[2] = userList[k].Email;
                    dt.Rows.Add(dr);
                }
            }

            DataView dv = new DataView(dt);
            tb1.DataSource = dv;
            tb1.DataBind();
            tb1.Width = 500;
            tb1.Height = 10;
            this.Controls.Add(tb1);

            Label outputFive = new Label();
            outputFive.Text = "<br /><br />";
            this.Controls.Add(outputFive);
            //Create a grid for groups
            tb2 = new DataGrid();
            tb2.AutoGenerateColumns = false;
            list = new CheckBoxList();

            //Create a new templatecolumn object
            TemplateColumn tcol2 = new TemplateColumn();
            {
                tcol2.HeaderText = "Select";
                tcol2.ItemTemplate = new DynamicItemTemplate_3();
            }

            DynamicItemTemplate_3 copy2 = (DynamicItemTemplate_3)tcol2.ItemTemplate;
            copy2.SetControls(this.Controls);
            copy2.SetCheckBoxes(list);
            allGroupMaps = new List<Map>();
            copy2.SetMap(allGroupMaps);

            tb2.Columns.Add(tcol2);
            tb2.Columns.Add(CreateBoundColumn("StringValue", "Group"));

            if (!Page.IsPostBack)
            {
                tb2.DataBind();
            }

            //Set up data
            dt2 = new DataTable();
            dt2.Columns.Add(new DataColumn("IntegerValue", tmpField.GetType()));
            dt2.Columns.Add(new DataColumn("StringValue", "".GetType()));

            if (groupList.Count > 0)
            {
                for (k = 0; k < groupList.Count; k++)
                {
                    dr = dt2.NewRow();
                    CheckBoxField field = new CheckBoxField();
                    dr[0] = field;
                    dr[1] = groupList[k].Name;
                    dt2.Rows.Add(dr);
                }
            }

            DataView dv2 = new DataView(dt2);
            tb2.DataSource = dv2;
            tb2.DataBind();
            tb2.Width = 500;
            String tb2Height = dt2.Rows.Count.ToString();
            tb2.Style.Add("Height", tb2Height);
            this.Controls.Add(tb2);

            Label outputSix = new Label();
            outputSix.Text = "<br /><br />Enter a message to send:<br /><br />";
            this.Controls.Add(outputSix);

            //Create controls for sending the email
            emailSubject = new TextBox();
            emailBody = new TextBox();
            emailSubject.Text = "- Email Subject -";
            emailSubject.TextMode = TextBoxMode.MultiLine;
            emailSubject.Width = 500;

            emailBody.Text = "- Email Body -";
            emailBody.TextMode = TextBoxMode.MultiLine;
            emailBody.Width = 500;
            emailBody.Rows = 6;
            this.Controls.Add(emailSubject);
            Label outputNine = new Label();
            outputNine.Text = "<br />";
            this.Controls.Add(outputNine);
            this.Controls.Add(emailBody);

            Label outputTen = new Label();
            outputTen.Text = "<br />";
            this.Controls.Add(outputTen);

            //Create control to add additional recipients to the email
            addRecipients = new TextBox();
            addRecipients.Text = "- Additional Recipients -";
            addRecipients.TextMode = TextBoxMode.MultiLine;
            addRecipients.Width = 500;
            addRecipients.Rows = 6;
            this.Controls.Add(addRecipients);

            Button finalizeForm = new Button();
            finalizeForm.Text = "Send Surveys";
            finalizeForm.Click += new EventHandler(sendSurveys_Click);

            Label outputEight = new Label();
            outputEight.Text = "<br /><br />";
            this.Controls.Add(outputEight);
            this.Controls.Add(finalizeForm);

        }
        private void LoadBoxData()
        {
            //Fills the drop down box with items containing "Survey" or "survey"
            Int16 index;
            SPListCollection collListTemp = SPContext.Current.Web.Lists;
            //Loop through elements of the collection, adding them all to the dropdownlist
            SPList currList;
            ListItem newItem = new ListItem();
            newItem.Text = "- Select Survey -";
            newItem.Value = 0.ToString();
            _sourceDropDownList.Items.Add(newItem);

            if (collListTemp.Count > 0)
            {
                for (index = 0; index < collListTemp.Count; index++)
                {
                    currList = collListTemp[index];
                    newItem = new ListItem();
                    //Set the new item's Text and Value
                    if (collListTemp[index].Title.Contains("Survey") ||
                        collListTemp[index].Title.Contains("survey"))
                    {
                        //Text is in the title, value is in the index in the collection
                        newItem.Text = collListTemp[index].Title;
                        newItem.Value = index.ToString();
                        _sourceDropDownList.Items.Add(newItem);
                    }
                }
            }
        }
        public void sendSurveys_Click(object sender, EventArgs e)
        {
            //Accesses user-selected data and submits the email.
            Label output = new Label();
            List<SPUser> finalList = new List<SPUser>();
            Label finalize = new Label();
            List<String> allEmails = new List<String>();
            String msgTo = "Emails: ";
            String surveyURL = "";
            Boolean send = true; //Change upon error
            Label errorLabel = new Label();
            errorLabel.Text = "";

            //Verifies if the user has provided valid input
            selectedSurvey = _sourceDropDownList.Items[_sourceDropDownList.SelectedIndex];
            if (selectedSurvey.Text == "- Select Survey -")
            {
                errorLabel.Text = "<br />Please choose valid selections.";
                send = false;
            }
            if (emailSubject.Text == "- Email Subject -")
            {
                errorLabel.Text = "<br />Please choose valid selections.";
                send = false;
            }
            if (emailBody.Text == "- Email Body -")
            {
                errorLabel.Text = "<br />Please choose valid selections.";
                send = false;
            }
            if (VerifyFormat() == false)
            {
                errorLabel.Text = "<br />Please provide a valid format for additional recipients. Separate emails with either a single space or a single comma.";
                send = false;
            }

            if (send == true)
            {
                //Correct format, proceed
                //Construct the allEmails list, adding emails from users and groups selected.
                Int16 i, j;
                for (i = 0; i < allMaps.Count; i++)
                {
                    if (allMaps[i].GetCheckBox().Checked == true)
                    {
                        //Check if already added
                        if (!allEmails.Contains(allMaps[i].GetUser().Email))
                        {
                            allEmails.Add(allMaps[i].GetUser().Email);
                        }
                    }
                }
                for (i = 0; i < allGroupMaps.Count; i++)
                {
                    if (allGroupMaps[i].GetCheckBox().Checked == true)
                    {
                        for (j = 0; j < allGroupMaps[i].GetGroup().Users.Count; j++)
                        {
                            //Check if the user was added
                            if (!allEmails.Contains(allGroupMaps[i].GetGroup().Users[j].Email))
                            {
                                allEmails.Add(allGroupMaps[i].GetGroup().Users[j].Email);
                            }
                        }
                    }
                }
                for (i = 0; i < addedEmails.Count; i++)
                {
                    //check if the user was added
                    if (!allEmails.Contains(addedEmails[i]))
                    {
                        allEmails.Add(addedEmails[i]);
                    }
                }
                //Construct the "msgTo" String which will be the to field of the email
                for (i = 0; i < allEmails.Count; i++)
                {
                    msgTo += allEmails[i];
                    if (!(i == (allEmails.Count - 1)))
                    {
                        msgTo += "; ";
                    }
                }
                //Provides a notice regarding the to field.
                Label toField = new Label();
                toField.Text = "<br /><br /><b>Note: The 'to' field of the email is currently the following:<br />";
                toField.Text += msgTo + "</b>";

                this.Controls.Add(toField);
                //Finds the URL of the selected survey
                Label showTest = new Label();

                //Send the email containing a link to the selected survey.
                SendEmail(emailSubject.Text, allEmails, emailBody.Text, allEmails.Count);
            }
            this.Controls.Add(errorLabel);
        }
        public void SendEmail(String subject, List<String> sendTo, String body, int numUsers)
        {
            //Sends the email using the SharePoint server's SMTP output email configuration. 
            Boolean passed = false;
            String surveyURL = "";
            String toField = "";
            Label emailSuccess, emailFailure;
            int i;
            emailSuccess = new Label();
            emailFailure = new Label();
            emailSuccess.Text = "";
            emailFailure.Text = "";


            surveyURL += SPContext.Current.Site.HostName;
            surveyURL += SPContext.Current.Web.Lists[selectedSurvey.ToString()].DefaultViewUrl;
            //Attempt to send an email via SMTP
            try
            {
                String replyTo = SPContext.Current.Site.WebApplication.OutboundMailReplyToAddress;
                String sender = SPContext.Current.Site.WebApplication.OutboundMailSenderAddress;
                String smtpAddress = SPContext.Current.Site.WebApplication.OutboundMailServiceInstance.Server.Address;
                MailMessage message;
                SmtpClient smtp;

                String singleToFieldString = "";
                for (i = 0; i < sendTo.Count; i++)
                {
                    singleToFieldString += sendTo[i];
                    if (i != (sendTo.Count - 1))
                    {
                        singleToFieldString += "; ";
                    }
                }
                toField = singleToFieldString;
                message = new MailMessage(sender, singleToFieldString, subject, body + "\"" + "file:///" + surveyURL + "\"");

                smtp = new SmtpClient(smtpAddress);
                smtp.Send(message);
                passed = true;
            }
            catch 
            {

            }
            if (passed == false)
            {
                try
                {
                    //Attempt to send the email via SPUtility
                    passed = SPUtility.SendEmail(SPContext.Current.Web, false, false, toField, subject, body + "\"" + "file:///" + surveyURL + "\"");
                }
                catch
                {

                }
            }

            if (passed == true)
            {
                emailSuccess.Text += "<br /><b>Email sent successfully.</b>";
            }
            else
            {
                emailFailure.Text += "<br /><br /><b>Cannot send email. The site's SMTP outbound email settings need to be set.</b>";
                ListItem selectionMade = _sourceDropDownList.Items[_sourceDropDownList.SelectedIndex];
                String surveyAlt = "";
                surveyAlt += SPContext.Current.Site.HostName;
                surveyAlt += SPContext.Current.Web.Lists[selectionMade.ToString()].DefaultViewUrl;
                emailFailure.Text += "<br /><br /><b>Note: </b>Alternatively, you can paste the following URL to the survey into an email, ";
                emailFailure.Text += "using the emails listed above.<br />";
                emailFailure.Text += "file:///" + surveyURL + "\"";
                emailFailure.Text += "<br /><br />";
            }
            this.Controls.Add(emailFailure);
            this.Controls.Add(emailSuccess);
        }

        public Boolean VerifyFormat()
        {
            //Checks the format of the additional recipients TextBox
            if (addRecipients.Text == "- Additional Recipients -")
            {
                return true;
            }
            if (addRecipients.Text == "")
            {
                return true;
            }

            //We have to manipulate the String and create a list of Strings
            String emailsString = addRecipients.Text;
            List<String> emails = new List<String>();
            Boolean flag = false;

            //Format: emails separated by spaces
            Boolean done = false;
            String[] arrStrings;
            Char[] splitChars = { ' ', ',' };
            String subString = addRecipients.Text.Trim();
            arrStrings = subString.Split(splitChars);

            //Remove extra white space
            int i, j;
            for (i = 0; i < arrStrings.Length; i++)
            {
                arrStrings[i] = arrStrings[i].Replace(" ", String.Empty);
            }

            for (i = 0; i < arrStrings.Length; i++)
            {
                //Get "name@ex.com"
                Regex emailRegEx = new Regex(@"^(?("")("".+?(?<!\\)""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z])@))" +
                @"(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-\w]*[0-9a-z]*\.)+[a-z0-9][\-a-z0-9]{0,22}[a-z0-9]))$");
                Match m = emailRegEx.Match(arrStrings[i]);
                if (m.Success)
                {
                    //Add the email address to the list
                    emails.Add(arrStrings[i].Trim());
                }
                else
                {
                    return false;
                }
            }
            SetEmailList(emails);
            return true;
        }
        public void SetEmailList(List<String> emails)
        {
            addedEmails = emails;
        }
    }

    public class Map
    {
        //This class maps a user to a checkbox object
        public SPUser _user = null;
        public SPGroup _group = null;
        public CheckBox _checkBox = null;
        public static int index = 0;
        public int pos = 0;
        public Boolean val = false;

        public void AddUser(SPUser user)
        {
            if (_user == null)
            {
                _user = user;
                pos = index;
                index = index + 1;
            }
        }
        public void AddGroup(SPGroup group)
        {
            if (_group == null)
            {
                _group = group;
                pos = index;
                index = index + 1;
            }
        }
        public void SetChoice(Boolean choice)
        {
            val = choice;
        }
        public void AddCheckBox(CheckBox checkbox)
        {
            if (_checkBox == null)
            {
                _checkBox = checkbox;
            }
        }
        public void ResetIndex()
        {
            index = 0;
        }
        public String GetUserName()
        {
            return _user.Name;
        }
        public String GetUserEmail()
        {
            return _user.Email;
        }
        public SPUser GetUser()
        {
            return _user;
        }
        public SPGroup GetGroup()
        {
            return _group;
        }
        public String GetGroupName()
        {
            return _group.Name;
        }
        public CheckBox GetCheckBox()
        {
            return _checkBox;
        }
        public String GetCheckBoxID()
        {
            return _checkBox.ID;
        }
        public Boolean GetChoice()
        {
            return val;
        }
    }
    public class DynamicItemTemplate_2 : ITemplate
    {
        //Template column for a user
        public ControlCollection controls;
        public CheckBoxList checkBoxes;
        public List<Map> maps = null;
        public SPUserCollection userList;
        public SPGroupCollection groupList;
        public int index = 0;

        public void InstantiateIn(Control container)
        {
            userList = SPContext.Current.Web.AllUsers;
            groupList = SPContext.Current.Web.Groups;
            Map newMap = new Map();
            newMap.AddUser(userList[index]);

            //Create an instance of a checkbox object.
            CheckBox oCheckBox = new CheckBox();
            oCheckBox.ID = "box_" + index.ToString();
            index += 1;
            oCheckBox.Checked = false;
            newMap.AddCheckBox(oCheckBox);

            Boolean exists = false;
            int i;
            for (i = 0; i < maps.Count; i++)
            {
                if (maps[i].GetUserName() == newMap.GetUserName())
                {
                    exists = true;
                    break;
                }
            }
            if (exists == false)
            {
                maps.Add(newMap);
            }

            //When the databinding event fires, properly bind.
            oCheckBox.CheckedChanged += new EventHandler(BindCheckBox);
            container.Controls.Add(oCheckBox);
        }
        public void SetUserList(SPUserCollection list)
        {
            userList = list;
        }
        public void SetControls(ControlCollection contr)
        {
            controls = contr;
        }
        public void SetCheckBoxes(CheckBoxList list)
        {
            checkBoxes = list;
        }
        public void IncreaseCounter()
        {
            index += 1;
        }
        public void DecreaseCounter()
        {
            index -= 1;
        }
        public void BindCheckBox(object sender, EventArgs e)
        {
            CheckBox oCheckBox = (CheckBox)sender;
            DataGridItem container = (DataGridItem)oCheckBox.NamingContainer;

            //Evaluate the data from the Grid item and set the Checked property 
            Label lab = new Label();
            int index, i;
            for (i = 0; i < maps.Count; i++)
            {
                if (maps[i].GetCheckBoxID() == oCheckBox.ID)
                {
                    index = i;
                    maps[i].SetChoice(oCheckBox.Checked);
                }
            }
        }
        public void SetMap(List<Map> listPointer)
        {
            maps = listPointer;
        }
    }

    public class DynamicItemTemplate_3 : ITemplate
    {
        //Template column for a group
        public ControlCollection controls;
        public CheckBoxList checkBoxes;
        public List<Map> maps = null;
        public SPUserCollection userList;
        public SPGroupCollection groupList;
        public int index = 0;

        public void InstantiateIn(Control container)
        {
            userList = SPContext.Current.Web.AllUsers;
            groupList = SPContext.Current.Web.Groups;
            Map newMap = new Map();
            newMap.AddGroup(groupList[index]);

            //Create an instance of a checkbox object.
            CheckBox oCheckBox = new CheckBox();
            oCheckBox.ID = "box_" + index.ToString();
            index += 1;
            oCheckBox.Checked = false;
            newMap.AddCheckBox(oCheckBox);

            Boolean exists = false;
            int i;
            for (i = 0; i < maps.Count; i++)
            {
                if (maps[i].GetGroupName() == newMap.GetGroupName())
                {
                    exists = true;
                    break;
                }
            }
            if (exists == false)
            {
                maps.Add(newMap);
            }

            //When the databinding event fires, properly bind.
            oCheckBox.CheckedChanged += new EventHandler(BindCheckBox);
            container.Controls.Add(oCheckBox);
        }
        public void SetUserList(SPUserCollection list)
        {
            userList = list;
        }
        public void SetControls(ControlCollection contr)
        {
            controls = contr;
        }
        public void SetCheckBoxes(CheckBoxList list)
        {
            checkBoxes = list;
        }
        public void IncreaseCounter()
        {
            index += 1;
        }
        public void DecreaseCounter()
        {
            index -= 1;
        }
        public void BindCheckBox(object sender, EventArgs e)
        {
            CheckBox oCheckBox = (CheckBox)sender;
            DataGridItem container = (DataGridItem)oCheckBox.NamingContainer;

            //Evaluate the data from the Grid item and set the Checked property 
            Label lab = new Label();
            int index, i;
            for (i = 0; i < maps.Count; i++)
            {
                if (maps[i].GetCheckBoxID() == oCheckBox.ID)
                {
                    index = i;
                    maps[i].SetChoice(oCheckBox.Checked);
                }
            }
        }
        public void SetMap(List<Map> listPointer)
        {
            maps = listPointer;
        }
    }

}
