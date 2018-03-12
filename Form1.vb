Option Explicit On
Imports System.Net.Mail
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Drawing
Imports GlobalLibrary
Imports System.Data.SqlClient
Imports Microsoft.Office.Interop



Public Class Form1

    Public DBParameters As New GlobalLibrary.DBParameters '(Enums.DatabaseMode.Datarepository, "USE1B-SQL12")
    Private Functions As New GlobalLibrary.Functions
    Private SQLHelper As New GlobalLibrary.SqlHelper
    Private usrApplicationManagment As New GlobalLibrary.ApplicationAccess.DRUser()
    Private CN As String = "Server=USE1B-SQL12;Database=DRGJira;Integrated Security=SSPI;Connection Timeout=150"
    Private cnSQL As SqlClient.SqlConnection = New SqlClient.SqlConnection(CN)

    Private bInitial As Boolean = True
    Private bModelTypeMod As Boolean
    Private dsDatabases As DataSet
    Private dsDatabaselist As DataSet
    Private dsSoftware As DataSet
    Private dsSoftwareList As DataSet
    Private dsTools As DataSet
    Private dsToolsList As DataSet
    Private dsProduct As DataSet
    Private dsProductList As DataSet
    Private dsProjectList As DataSet
    Private dsPM As DataSet
    Private dsPMList As DataSet
    Private dsPMforProject As DataSet
    Private dsPO As DataSet
    Private dsPOList As DataSet
    Private dsPOforProject As DataSet
    Private dsResources As DataSet
    Private dsActiveTickets As DataSet
    Private dsMTDHours As DataSet
    Private dsProductTickets As DataSet
    Private dsGroupList As DataSet
    Private dsGroup As DataSet
    Private dsGroups As DataSet
    Private dsGroupDetails As DataSet
    Private dsEngineers As DataSet
    Private dsGroupMTD As DataSet, dsGroupPrior As DataSet, dsGroupThis As DataSet, dsProductMTD As DataSet, dsProductPrior As DataSet
    Private dsGroupActiveTickets As DataSet, dsGroupTODOTickets As DataSet, dsGroupAssignments As DataSet
    Private dsFeatureList As DataSet, dsFeature As DataSet, dsfeatureitem As DataSet, dsFeatureManagement As DataSet, dsFeatureProject As DataSet, dsMonthPicker As DataSet, dsFeatureResource As DataSet
    Private dsfeatureproduct As DataSet, dsfeaturebreakout As DataSet, dsfeaturetime As DataSet, dsmissingtickets As DataSet
    Public Userid As String
    'Public BindingSourceL, BindingSource1, BindingSource2, BindingSource3 As BindingSource
    Dim myFormat As String = "HH:mm"


    Private Sub Form1_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Try
            Userid = ApplicationAccess.DRUser.UserID

            dsDatabases = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_database")
            cmbDatabases.DataSource = dsDatabases.Tables(0)
            cmbDatabases.DisplayMember = dsDatabases.Tables(0).Columns("Description").ToString

            dsDatabaselist = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_databaseList")

            dsSoftware = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_software")
            cmbsoftware.DataSource = dsSoftware.Tables(0)
            cmbsoftware.DisplayMember = dsSoftware.Tables(0).Columns("Description").ToString

            dsSoftwareList = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_SoftwareList")

            dsTools = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_Tools")
            cmbTools.DataSource = dsTools.Tables(0)
            cmbTools.DisplayMember = dsTools.Tables(0).Columns("Description").ToString

            dsToolsList = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_ToolsList")

            dsProductList = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_ProductList")
            cmbProduct.DataSource = dsProductList.Tables(0)
            cmbProduct.DisplayMember = dsProductList.Tables(0).Columns("Description").ToString

            dsGroupList = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_GroupList")
            cmbGroups.DataSource = dsGroupList.Tables(0)
            cmbGroups.DisplayMember = dsGroupList.Tables(0).Columns("Description").ToString

            dsGroups = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_GroupList")
            cmbGroup.DataSource = dsGroups.Tables(0)
            cmbGroup.DisplayMember = dsGroups.Tables(0).Columns("Description").ToString

            dsProjectList = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_ProjectList")
            ComboBox2.DataSource = dsProjectList.Tables(0)
            ComboBox2.DisplayMember = dsProjectList.Tables(0).Columns("Name").ToString

            dsFeatureProject = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_ProjectList")
            ComboBox6.DataSource = dsFeatureProject.Tables(0)
            ComboBox6.DisplayMember = dsFeatureProject.Tables(0).Columns("Name").ToString

            dsPMList = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_PMList")
            cmbPM.DataSource = dsPMList.Tables(0)
            cmbPM.DisplayMember = dsPMList.Tables(0).Columns("Name").ToString

            dsPMforProject = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_PMList")
            cmbPMList.DataSource = dsPMforProject.Tables(0)
            cmbPMList.DisplayMember = dsPMforProject.Tables(0).Columns("Name").ToString

            dsPOList = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_POList")
            cmbProjectOwner.DataSource = dsPOList.Tables(0)
            cmbProjectOwner.DisplayMember = dsPOList.Tables(0).Columns("Name").ToString

            dsPOforProject = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_POList")
            cmbPOList.DataSource = dsPOforProject.Tables(0)
            cmbPOList.DisplayMember = dsPOforProject.Tables(0).Columns("Name").ToString

            dsResources = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_ResourceList", IIf(ckbEngineers.CheckState = CheckState.Checked, 1, 0))
            cmbResources.DataSource = dsResources.Tables(0)
            cmbResources.DisplayMember = dsResources.Tables(0).Columns("Resource").ToString  '+ " " + dsResources.Tables(0).Columns("lastname").ToString

            dsActiveTickets = SQLHelper.ExecuteDataset(CN, "dbo.s_get_assigned_tickets", 0)
            dgvActiveTickets_FormatGrid()
            dgvActiveTickets_BindData()

            dsMTDHours = SQLHelper.ExecuteDataset(CN, "dbo.s_Month_to_Date_hours_by_Product", 0)
            dgvMTDHours_FormatGrid()
            dgvMTDHours_BindData()

            dsGroupPrior = SQLHelper.ExecuteDataset(CN, "dbo.s_get_group_lastweek_hours", 0)
            dgvPriorGroupHours_FormatGrid()
            dgvPriorGroupHours_BindData()

            dsGroupThis = SQLHelper.ExecuteDataset(CN, "dbo.s_get_group_thisweek_hours", 0)
            dgvThisGroupHours_FormatGrid()
            dgvThisGroupHours_BindData()

            dsProductTickets = SQLHelper.ExecuteDataset(CN, "dbo.s_get_products_for_Products_tab", "")
            dgvProductsTickets_FormatGrid()
            dgvProductsTickets_BindData()

            dsEngineers = SQLHelper.ExecuteDataset(CN, "dbo.s_get_Engineers")

            cmbGroupLead.BindingContext = New BindingContext
            cmbGroupLead.DataSource = dsEngineers.Tables(0)
            cmbGroupLead.DisplayMember = dsEngineers.Tables(0).Columns("Name").ToString
            cmbGroupLead.ValueMember = dsEngineers.Tables(0).Columns("ResourceID").ToString

            cmbGroupDB1.BindingContext = New BindingContext
            cmbGroupDB1.DataSource = dsEngineers.Tables(0)
            cmbGroupDB1.DisplayMember = dsEngineers.Tables(0).Columns("Name").ToString
            cmbGroupDB1.ValueMember = dsEngineers.Tables(0).Columns("ResourceID").ToString

            cmbGroupDB2.BindingContext = New BindingContext
            cmbGroupDB2.DataSource = dsEngineers.Tables(0)
            cmbGroupDB2.DisplayMember = dsEngineers.Tables(0).Columns("Name").ToString
            cmbGroupDB2.ValueMember = dsEngineers.Tables(0).Columns("ResourceID").ToString

            cmbGroupDB3.BindingContext = New BindingContext
            cmbGroupDB3.DataSource = dsEngineers.Tables(0)
            cmbGroupDB3.DisplayMember = dsEngineers.Tables(0).Columns("Name").ToString
            cmbGroupDB3.ValueMember = dsEngineers.Tables(0).Columns("ResourceID").ToString

            dsGroupMTD = SQLHelper.ExecuteDataset(CN, "dbo.s_get_group_monthtodate_hours", 0)
            dgvMTDGroupHours_FormatGrid()
            dgvMTDGroupHours_BindData()

            dsProductMTD = SQLHelper.ExecuteDataset(CN, "dbo.s_get_group_product_MTD_hours", 0)
            dgvProjectMTD_FormatGrid()
            dgvProjectMTD_BindData()

            dsProductPrior = SQLHelper.ExecuteDataset(CN, "dbo.s_get_group_product_Prior_hours", 0)
            dgvProjectPrior_FormatGrid()
            dgvProjectPrior_BindData()

            dsGroupActiveTickets = SQLHelper.ExecuteDataset(CN, "dbo.s_get_Group_Active_Tickets", 0)
            dgvGroupActiveTickets_FormatGrid()
            dgvGroupActiveTickets_BindData()

            dsGroupTODOTickets = SQLHelper.ExecuteDataset(CN, "dbo.s_get_Group_TODO_Tickets", 0)
            dgvTODOTickets_FormatGrid()
            dgvTODOTickets_BindData()

            dsFeatureList = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_FeatureList", 0)
            cmbfeature.DataSource = dsFeatureList.Tables(0)
            cmbfeature.DisplayMember = dsFeatureList.Tables(0).Columns("Description").ToString


            dsFeatureResource = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_months")
            ComboBox4.DataSource = dsFeatureResource.Tables(0)
            ComboBox4.DisplayMember = dsFeatureResource.Tables(0).Columns("Date").ToString

            dsMonthPicker = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_Engineers")
            ComboBox5.DataSource = dsMonthPicker.Tables(0)
            ComboBox5.DisplayMember = dsMonthPicker.Tables(0).Columns("Name").ToString

            'dsFeature = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_FeatureList", 1)
            'cmbfeatureMgmt.DataSource = dsFeature.Tables(0)
            'cmbfeatureMgmt.DisplayMember = dsFeature.Tables(0).Columns("Description").ToString


            'dsProductList = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_ProductList")

            cbPMDB1.Visible = False
            cbPMDB2.Visible = False
            cbPMDB3.Visible = False
            cbPMDB4.Visible = False
            cbPMSW1.Visible = False
            cbPMSW2.Visible = False
            cbPMSW3.Visible = False
            cbPMSW4.Visible = False
            cbPMTL1.Visible = False
            cbPMTL2.Visible = False
            cbPMTL3.Visible = False
            cbPMTL4.Visible = False
            'cbPMPrd1.Visible = False
            'cbPMPrd2.Visible = False
            'cbPMPrd3.Visible = False
            'cbPMPrd4.Visible = False
            'cbPMPrd5.Visible = False
            'cbPMPrd6.Visible = False
            'cbPMPrd7.Visible = False
            'cbPMPrd8.Visible = False

            If dsDatabaselist.Tables(0).Rows.Count > 0 Then
                cbPMDB1.Text = dsDatabaselist.Tables(0).Rows(0).Item(1)
                cbPMDB1.Visible = True
                If dsDatabaselist.Tables(0).Rows.Count > 1 Then
                    cbPMDB2.Text = dsDatabaselist.Tables(0).Rows(1).Item(1)
                    cbPMDB2.Visible = True
                End If
                If dsDatabaselist.Tables(0).Rows.Count > 2 Then
                    cbPMDB3.Text = dsDatabaselist.Tables(0).Rows(2).Item(1)
                    cbPMDB3.Visible = True
                End If
                If dsDatabaselist.Tables(0).Rows.Count > 3 Then
                    cbPMDB4.Text = dsDatabaselist.Tables(0).Rows(3).Item(1)
                    cbPMDB4.Visible = True
                End If
            End If


            If dsSoftwareList.Tables(0).Rows.Count > 0 Then
                cbPMSW1.Text = dsSoftwareList.Tables(0).Rows(0).Item(1)
                cbPMSW1.Visible = True
                If dsSoftwareList.Tables(0).Rows.Count > 1 Then
                    cbPMSW2.Text = dsSoftwareList.Tables(0).Rows(1).Item(1)
                    cbPMSW2.Visible = True
                End If
                If dsSoftwareList.Tables(0).Rows.Count > 2 Then
                    cbPMSW3.Text = dsSoftwareList.Tables(0).Rows(2).Item(1)
                    cbPMSW3.Visible = True
                End If
                If dsSoftwareList.Tables(0).Rows.Count > 3 Then
                    cbPMSW4.Text = dsSoftwareList.Tables(0).Rows(3).Item(1)
                    cbPMSW4.Visible = True
                End If
            End If

            If dsToolsList.Tables(0).Rows.Count > 0 Then
                cbPMTL1.Text = dsToolsList.Tables(0).Rows(0).Item(1)
                cbPMTL1.Visible = True
                If dsToolsList.Tables(0).Rows.Count > 1 Then
                    cbPMTL2.Text = dsToolsList.Tables(0).Rows(1).Item(1)
                    cbPMTL2.Visible = True
                End If
                If dsToolsList.Tables(0).Rows.Count > 2 Then
                    cbPMTL3.Text = dsToolsList.Tables(0).Rows(2).Item(1)
                    cbPMTL3.Visible = True
                End If
                If dsToolsList.Tables(0).Rows.Count > 3 Then
                    cbPMTL4.Text = dsToolsList.Tables(0).Rows(3).Item(1)
                    cbPMTL4.Visible = True
                End If
            End If

            'If dsProductList.Tables(0).Rows.Count > 0 Then
            '    cbPMPrd1.Text = dsProductList.Tables(0).Rows(0).Item(1)
            '    cbPMPrd1.Visible = True
            '    If dsProductList.Tables(0).Rows.Count > 1 Then
            '        cbPMPrd2.Text = dsProductList.Tables(0).Rows(1).Item(1)
            '        cbPMPrd2.Visible = True
            '    End If
            '    If dsProductList.Tables(0).Rows.Count > 2 Then
            '        cbPMPrd3.Text = dsProductList.Tables(0).Rows(2).Item(1)
            '        cbPMPrd3.Visible = True
            '    End If
            '    If dsProductList.Tables(0).Rows.Count > 3 Then
            '        cbPMPrd4.Text = dsProductList.Tables(0).Rows(3).Item(1)
            '        cbPMPrd4.Visible = True
            '    End If
            '    If dsProductList.Tables(0).Rows.Count > 4 Then
            '        cbPMPrd5.Text = dsProductList.Tables(0).Rows(4).Item(1)
            '        cbPMPrd5.Visible = True
            '    End If
            '    If dsProductList.Tables(0).Rows.Count > 5 Then
            '        cbPMPrd6.Text = dsProductList.Tables(0).Rows(5).Item(1)
            '        cbPMPrd6.Visible = True
            '    End If
            '    If dsProductList.Tables(0).Rows.Count > 6 Then
            '        cbPMPrd7.Text = dsProductList.Tables(0).Rows(6).Item(1)
            '        cbPMPrd7.Visible = True
            '    End If
            '    If dsProductList.Tables(0).Rows.Count > 7 Then
            '        cbPMPrd8.Text = dsProductList.Tables(0).Rows(7).Item(1)
            '        cbPMPrd8.Visible = True
            '    End If
            'End If
            dsGroupAssignments = SQLHelper.ExecuteDataset(CN, "dbo.s_get_assigned_Projects", 0)


            dgvGroupAssignments_FormatGrid()
            dgvGroupAssignments_BindData()

            bInitial = False

        Catch ex As Exception
            'Functions.Sendmail(ex.Message, "Form Load" + " : " + Userid, 0, 0, "btnludatabasesNew_Click")
            MsgBox(ex.Message)
        End Try

    End Sub


#Region "Combo boxes"
    Private Sub cmbDatabases_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbDatabases.SelectedIndexChanged
        Try
            If Not bInitial Then
                bInitial = True

                txtdatabasesNamedesc.Text = cmbDatabases.SelectedValue(1).ToString

                btnludatabasesSave.Visible = False
                btnludatabasesCancel.Visible = False
                bInitial = False
            End If

        Catch ex As Exception
            MsgBox("cmbDatabases_SelectedIndexChanged" + " : " + ex.Message)
            'Functions.Sendmail(ex.Message, "cmbDatabases_SelectedIndexChanged" + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub

    Private Sub cmbsoftware_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbSoftware.SelectedIndexChanged
        Try
            If Not bInitial Then
                bInitial = True

                txtsoftwareNamedesc.Text = cmbSoftware.SelectedValue(1).ToString

                btnlusoftwareSave.Visible = False
                btnluSoftwareCancel.Visible = False
                bInitial = False
                txtSoftwareNamedesc.ReadOnly = False
            End If

        Catch ex As Exception
            'Functions.Sendmail(ex.Message, "cmbsoftware_SelectedIndexChanged" + " : " + Userid, 0, 0, "Project Management")
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub cmbTools_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbTools.SelectedIndexChanged
        Try
            If Not bInitial Then
                bInitial = True
                txtToolsNamedesc.Text = cmbTools.SelectedValue(1).ToString
                btnluToolsSave.Visible = False
                btnluToolsCancel.Visible = False
                bInitial = False
                txtToolsNamedesc.ReadOnly = False
            End If

        Catch ex As Exception
            MsgBox("cmbTools_SelectedIndexChanged" + " : " + ex.Message)
            'Functions.Sendmail(ex.Message, "cmbTools_SelectedIndexChanged" + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub

    Private Sub cmbProduct_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbProduct.SelectedIndexChanged
        Dim myindex As Integer
        Try
            If Not bInitial Then
                bInitial = True
                myindex = cmbProduct.SelectedIndex
                dsProduct = SQLHelper.ExecuteDataset(CN, "dbo.s_get_Product", cmbProduct.SelectedValue(0))

                txtProductNamedesc.Text = isnull(dsProduct.Tables(0).Rows(0).Item("Description"))
                Functions.whyareyousodumb(CheckBox29, dsProduct.Tables(0).Rows(0).Item("isActive"))
                txtProductNamedesc.ReadOnly = False
                btnluProductSave.Visible = False
                btnluProductCancel.Visible = False
                CheckBox29.Enabled = True
                bInitial = False
                cmbProduct.SelectedIndex = myindex
            End If

        Catch ex As Exception
            MsgBox("cmbProduct_SelectedIndexChanged" + " : " + ex.Message)
            'Functions.Sendmail(ex.Message, "cmbProduct_SelectedIndexChanged" + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub

    Private Sub cmbGroup_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbGroup.SelectedIndexChanged
        Dim myindex As Integer
        Try
            If Not bInitial Then
                bInitial = True
                myindex = cmbGroup.SelectedIndex
                If cmbGroup.SelectedIndex <> 0 Then
                    dsGroup = SQLHelper.ExecuteDataset(CN, "dbo.s_get_Group", cmbGroup.SelectedValue(0))
                    txtGroupDesc.Text = isnull(dsGroup.Tables(0).Rows(0).Item("Description"))
                    Functions.whyareyousodumb(ckbGroupActive, dsGroup.Tables(0).Rows(0).Item("isActive"))
                Else
                    txtGroupDesc.Clear()
                    ckbGroupActive.CheckState = CheckState.Indeterminate
                    txtGroupDesc.ReadOnly = True
                    ckbGroupActive.Enabled = False
                End If
                txtGroupDesc.ReadOnly = False
                btnSaveGroup.Visible = False
                btnCancelGroup.Visible = False
                ckbGroupActive.Enabled = True
                cmbGroup.SelectedIndex = myindex
                bInitial = False
            End If

        Catch ex As Exception
            MsgBox("cmbGroup_SelectedIndexChanged" + " : " + ex.Message)
            'Functions.Sendmail(ex.Message, "cmbGroup_SelectedIndexChanged" + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub

    Private Sub cmbResources_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbResources.SelectedIndexChanged
        Dim myDS As DataSet
        Try
            If Not bInitial Then
                bInitial = True
                If cmbResources.SelectedValue(0) <> 0 Then
                    myDS = SQLHelper.ExecuteDataset(CN, "dbo.s_get_Resource", cmbResources.SelectedValue(0))
                    ' load screen
                    txtResource.Text = isnull(myDS.Tables(0).Rows(0).Item("Resource"))
                    txtResourceFirstName.Text = isnull(myDS.Tables(0).Rows(0).Item("FirstName"))
                    txtResourceLastName.Text = isnull(myDS.Tables(0).Rows(0).Item("LastName"))
                    txtResourceRole.Text = isnull(myDS.Tables(0).Rows(0).Item("Role"))
                    txtResourceTimezone.Text = isnull(myDS.Tables(0).Rows(0).Item("Timezone"))
                    txtHrsMTD.Text = isnull(myDS.Tables(0).Rows(0).Item("HoursMTD"))
                    txtCapExMTD.Text = isnull(myDS.Tables(0).Rows(0).Item("CapExMTD"))
                    txtPctCapEx.Text = FormatPercent(isNumNull(myDS.Tables(0).Rows(0).Item("pct_CapEx")))
                    rtbResourceNote.Text = isnull(myDS.Tables(0).Rows(0).Item("Note"))
                    Functions.whyareyousodumb(ckbResourceEngineer, myDS.Tables(0).Rows(0).Item("isengineer"))
                    Functions.whyareyousodumb(ckbResourceisActive, myDS.Tables(0).Rows(0).Item("isActive"))
                    dsActiveTickets = SQLHelper.ExecuteDataset(CN, "dbo.s_get_assigned_tickets", cmbResources.SelectedValue(0))
                    dsMTDHours = SQLHelper.ExecuteDataset(CN, "dbo.s_Month_to_Date_hours_by_Product", cmbResources.SelectedValue(0))
                    txtStartTime.Text = Format(CDate(isTimeNull(myDS.Tables(0).Rows(0).Item("StartTime").ToString)), "HH:mm")
                    txtEndTime.Text = Format(CDate(isTimeNull(myDS.Tables(0).Rows(0).Item("EndTime").ToString)), "HH:mm")
                    Label123.Text = cmbResources.SelectedValue(0)
                Else
                    txtResource.Clear()
                    txtResourceFirstName.Clear()
                    txtResourceLastName.Clear()
                    txtResourceRole.Clear()
                    txtResourceTimezone.Clear()
                    txtHrsMTD.Clear()
                    txtCapExMTD.Clear()
                    txtPctCapEx.Clear()
                    rtbResourceNote.Clear()
                    ckbResourceEngineer.CheckState = CheckState.Indeterminate
                    ckbResourceisActive.CheckState = CheckState.Indeterminate
                    dsActiveTickets.Clear()
                    dsMTDHours.Clear()
                    txtStartTime.Clear()
                    txtEndTime.Clear()
                    Label123.Text = ""
                End If

                dgvActiveTickets_BindData()
                dgvMTDHours_BindData()
                btnSaveResource.Visible = False
                btnCancelResource.Visible = False
                bModelTypeMod = True
                bInitial = False



            End If

        Catch ex As Exception
            MsgBox("cmbResources_SelectedIndexChanged" + " : " + ex.Message)
            'Functions.Sendmail(ex.Message, "cmbResources_SelectedIndexChanged" + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged

        Try
            If Not bInitial Then
                bInitial = True

                dsProduct = SQLHelper.ExecuteDataset(CN, "dbo.s_get_Project", ComboBox2.SelectedValue(0))
                Label91.Text = ComboBox2.SelectedValue(0)
                If dsProduct.Tables(0).Rows.Count > 0 Then
                    TextBox6.Text = isnull(dsProduct.Tables(0).Rows(0).Item("Name"))
                    '--servers
                    TextBox81.Text = isnull(dsProduct.Tables(0).Rows(0).Item("TestingDBServer"))
                    TextBox82.Text = isnull(dsProduct.Tables(0).Rows(0).Item("QADBServer"))
                    TextBox83.Text = isnull(dsProduct.Tables(0).Rows(0).Item("StagingDBServer"))
                    TextBox84.Text = isnull(dsProduct.Tables(0).Rows(0).Item("ProductionDBServer"))
                    txtPandL.Text = isnull(dsProduct.Tables(0).Rows(0).Item("PandL"))
                    TextBox2.Text = isnull(dsProduct.Tables(0).Rows(0).Item("DBLead"))
                    TextBox3.Text = isnull(dsProduct.Tables(0).Rows(0).Item("DBDev1"))
                    TextBox4.Text = isnull(dsProduct.Tables(0).Rows(0).Item("DBDev2"))
                    TextBox5.Text = isnull(dsProduct.Tables(0).Rows(0).Item("DBDev3"))
                    '-----------  boxes
                    TextBox7.Text = isnull(dsProduct.Tables(0).Rows(0).Item("ConfluencePage"))
                    TextBox85.Text = isnull(dsProduct.Tables(0).Rows(0).Item("Version"))
                    RichTextBox14.Text = isnull(dsProduct.Tables(0).Rows(0).Item("Notes"))
                    If isnull(dsProduct.Tables(0).Rows(0).Item("isActive")) = "" Then
                        ckbProjectActive.CheckState = CheckState.Indeterminate
                    Else
                        ckbProjectActive.CheckState = IIf(isnull(dsProduct.Tables(0).Rows(0).Item("isActive")) = False, CheckState.Unchecked, CheckState.Checked)
                    End If

                    If isnull(dsProduct.Tables(0).Rows(0).Item("isCapEx")) = "" Then
                        ckbCapEx.CheckState = CheckState.Indeterminate
                    Else
                        ckbCapEx.CheckState = IIf(isnull(dsProduct.Tables(0).Rows(0).Item("isCapEx")) = False, CheckState.Unchecked, CheckState.Checked)
                    End If

                    cmbPMList.SelectedIndex = 0

                    If dsProduct.Tables(0).Rows(0).Item("ProjectManagerID") = 0 Then
                        cmbPMList.SelectedIndex = 0
                        TextBox72.Clear()
                    Else
                        cmbPMList.SelectedText = dsProduct.Tables(0).Rows(0).Item("Project_manager")
                        TextBox72.Text = isnull(dsProduct.Tables(0).Rows(0).Item("PM_Location"))
                    End If

                    'cmbPOList.SelectedIndex = 0

                    If dsProduct.Tables(0).Rows(0).Item("ProjectOwnerID") = 0 Then
                        cmbPOList.SelectedIndex = 0
                        TextBox71.Clear()
                    Else
                        cmbPOList.SelectedText = dsProduct.Tables(0).Rows(0).Item("Project_owner")
                        TextBox71.Text = isnull(dsProduct.Tables(0).Rows(0).Item("PO_Location"))
                    End If

                    dsProductTickets = SQLHelper.ExecuteDataset(CN, "dbo.s_get_products_for_Products_tab", isnull(dsProduct.Tables(0).Rows(0).Item("Name")))
                    dgvProductsTickets_BindData()

                Else
                    Clear_project()
                End If

                bInitial = False
            End If

        Catch ex As Exception
            MsgBox("ComboBox2_SelectedIndexChanged" + " : " + ex.Message)
            ' Functions.Sendmail(ex.Message, "ComboBox2_SelectedIndexChanged" + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub

    Private Sub cmbPM_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbPM.SelectedIndexChanged
        'Dim myindex As Integer
        Try
            If Not bInitial Then
                bInitial = True
                ' myindex = cmbPM.SelectedIndex
                dsPM = SQLHelper.ExecuteDataset(CN, "dbo.s_get_PM", cmbPM.SelectedValue(0))

                If dsPM.Tables(0).Rows.Count = 0 Then
                    txtPMName.Clear()
                    txtPMLocation.Clear()
                    ckbPMActive.CheckState = CheckState.Unchecked
                    txtPMName.ReadOnly = True
                    txtPMLocation.ReadOnly = True
                    btnPMSave.Visible = False
                    btnPMCancel.Visible = False
                    ckbPMActive.Enabled = False
                Else
                    txtPMName.Text = isnull(dsPM.Tables(0).Rows(0).Item("Name"))
                    txtPMLocation.Text = isnull(dsPM.Tables(0).Rows(0).Item("Location"))
                    Functions.whyareyousodumb(ckbPMActive, dsPM.Tables(0).Rows(0).Item("isActive"))
                    txtPMName.ReadOnly = False
                    txtPMLocation.ReadOnly = False
                    btnPMSave.Visible = False
                    btnPMCancel.Visible = False
                    ckbPMActive.Enabled = True
                End If

                bInitial = False
                'cmbPM.SelectedIndex = myindex
            End If

        Catch ex As Exception
            MsgBox("cmbPM_SelectedIndexChanged" + " : " + ex.Message)
            'Functions.Sendmail(ex.Message, "cmbPM_SelectedIndexChanged" + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub

    Private Sub cmbPMList_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbPMList.SelectedIndexChanged
        Dim myDS As DataSet
        Try
            If Not bInitial Then
                bInitial = True
                myDS = SQLHelper.ExecuteDataset(CN, "dbo.s_get_PM", cmbPMList.SelectedValue(0))
                TextBox72.Text = isnull(myDS.Tables(0).Rows(0).Item("Location"))
                btnSaveProjectMaintenance.Visible = True
                btnCancelProjectMaintenance.Visible = True
                bModelTypeMod = True
                bInitial = False
            End If

        Catch ex As Exception
            MsgBox("cmbPMList_SelectedIndexChanged" + " : " + ex.Message)
            'Functions.Sendmail(ex.Message, "cmbPMList_SelectedIndexChanged" + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub

    Private Sub cmbPOList_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbPOList.SelectedIndexChanged
        Dim myDS As DataSet
        Try
            If Not bInitial Then
                bInitial = True
                myDS = SQLHelper.ExecuteDataset(CN, "dbo.s_get_PO", cmbPOList.SelectedValue(0))
                TextBox71.Text = isnull(myDS.Tables(0).Rows(0).Item("Location"))
                btnSaveProjectMaintenance.Visible = True
                btnCancelProjectMaintenance.Visible = True
                bModelTypeMod = True
                bInitial = False
            End If

        Catch ex As Exception
            MsgBox("cmbPOList_SelectedIndexChanged" + " : " + ex.Message)
            'Functions.Sendmail(ex.Message, "cmbPOList_SelectedIndexChanged" + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub

    Private Sub cmbProjectOwner_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbProjectOwner.SelectedIndexChanged
        'Dim myindex As Integer
        Try
            If Not bInitial Then
                bInitial = True
                'myindex = cmbProjectOwner.SelectedIndex
                dsPO = SQLHelper.ExecuteDataset(CN, "dbo.s_get_PO", cmbProjectOwner.SelectedValue(0))

                If dsPO.Tables(0).Rows.Count = 0 Then
                    txtPOName.Clear()
                    txtPOLocation.Clear()
                    ckbPOActive.CheckState = CheckState.Unchecked
                    txtPOName.ReadOnly = True
                    txtPOLocation.ReadOnly = True
                    btnPOSave.Visible = False
                    btnPOCancel.Visible = False
                    ckbPOActive.Enabled = False
                Else
                    txtPOName.Text = isnull(dsPO.Tables(0).Rows(0).Item("Name"))
                    txtPOLocation.Text = isnull(dsPO.Tables(0).Rows(0).Item("Location"))
                    Functions.whyareyousodumb(ckbPOActive, dsPO.Tables(0).Rows(0).Item("isActive"))
                    txtPOName.ReadOnly = False
                    txtPOLocation.ReadOnly = False
                    btnPOSave.Visible = False
                    btnPOCancel.Visible = False
                    ckbPOActive.Enabled = True
                End If
                bInitial = False
                'cmbProjectOwner.SelectedIndex = myindex
            End If

        Catch ex As Exception
            MsgBox("cmbPM_SelectedIndexChanged" + " : " + ex.Message)
            'Functions.Sendmail(ex.Message, "cmbPM_SelectedIndexChanged" + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub

    Private Sub cmbGroups_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbGroups.SelectedIndexChanged
        Try
            If Not bInitial Then
                bInitial = True
                Dim myi As Integer = cmbGroups.SelectedValue(0)
                If cmbGroups.SelectedValue(0) <> 0 Then
                    dsGroupMTD = SQLHelper.ExecuteDataset(CN, "dbo.s_get_group_monthtodate_hours", cmbGroups.SelectedValue(0))
                    dsGroupPrior = SQLHelper.ExecuteDataset(CN, "dbo.s_get_group_lastweek_hours", cmbGroups.SelectedValue(0))
                    dsGroupThis = SQLHelper.ExecuteDataset(CN, "dbo.s_get_group_Thisweek_hours", cmbGroups.SelectedValue(0))
                    dsProductMTD = SQLHelper.ExecuteDataset(CN, "dbo.s_get_group_product_MTD_hours", cmbGroups.SelectedValue(0))
                    dsProductPrior = SQLHelper.ExecuteDataset(CN, "dbo.s_get_group_product_Prior_hours", cmbGroups.SelectedValue(0))
                    dsGroupActiveTickets = SQLHelper.ExecuteDataset(CN, "dbo.s_get_Group_Active_Tickets", cmbGroups.SelectedValue(0))
                    dsGroupTODOTickets = SQLHelper.ExecuteDataset(CN, "dbo.s_get_Group_todo_Tickets", cmbGroups.SelectedValue(0))
                Else
                    dsGroupMTD.Clear()
                    dsGroupPrior.Clear()
                    dsGroupThis.Clear()
                    dsProductMTD.Clear()
                    dsProductPrior.Clear()
                    dsGroupActiveTickets.Clear()
                    dsGroupTODOTickets.Clear()
                End If
                dgvMTDGroupHours_BindData()
                dgvPriorGroupHours_BindData()
                dgvThisGroupHours_BindData()
                dgvProjectMTD_BindData()
                dgvProjectPrior_BindData()
                dgvGroupActiveTickets_BindData()
                dgvTODOTickets_BindData()

                dsGroupDetails = SQLHelper.ExecuteDataset(CN, "dbo.s_get_GroupDetails", cmbGroups.SelectedValue(0))
                If dsGroupDetails.Tables(0).Rows.Count > 0 Then
                    cmbGroupLead.SelectedValue = isNumNull(dsGroupDetails.Tables(0).Rows(0).Item("LeadID").ToString)
                    cmbGroupDB1.SelectedValue = isNumNull(dsGroupDetails.Tables(0).Rows(0).Item("DB1_ID").ToString)
                    cmbGroupDB2.SelectedValue = isNumNull(dsGroupDetails.Tables(0).Rows(0).Item("DB2_ID").ToString)
                    cmbGroupDB3.SelectedValue = isNumNull(dsGroupDetails.Tables(0).Rows(0).Item("DB3_ID").ToString)
                Else
                    cmbGroupLead.SelectedIndex = 0
                    cmbGroupDB1.SelectedIndex = 0
                    cmbGroupDB2.SelectedIndex = 0
                    cmbGroupDB3.SelectedIndex = 0
                End If

                dsGroupAssignments = SQLHelper.ExecuteDataset(CN, "dbo.s_get_assigned_Projects", cmbGroups.SelectedValue(0))
                dgvGroupAssignments_BindData()

                If dsGroupAssignments.Tables(0).Rows.Count > 0 Then
                    For i As Integer = 1 To dsGroupAssignments.Tables(0).Rows.Count - 1
                        If dsGroupAssignments.Tables(0).Rows(i).Item("isUsed") = 1 And dsGroupAssignments.Tables(0).Rows(i).Item("isAssigned") = 0 Then
                            ''clbAssignments.Items(i
                        End If
                    Next
                End If

                bInitial = False
            End If
        Catch ex As Exception
            MsgBox("cmbGroups_SelectedIndexChanged" + " : " + ex.Message)
            'Functions.Sendmail(ex.Message, "cmbGroups_SelectedIndexChanged" + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub


#End Region

#Region "buttons"
    Private Sub btnludatabasesNew_Click(sender As System.Object, e As System.EventArgs) Handles btnludatabasesNew.Click
        Try
            bInitial = True
            txtdatabasesNamedesc.Clear()
            bInitial = False
            cmbDatabases.SelectedIndex = 0
        Catch ex As Exception
            MsgBox("btnludatabasesNew_Click : New Database  : " + ex.Message)
            'Functions.Sendmail(ex.Message, "btnludatabasesNew_Click " + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub

    Private Sub btnludatabasesCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnludatabasesCancel.Click
        cmbDatabases_SelectedIndexChanged(cmbDatabases, EventArgs.Empty)
        txtdatabasesNamedesc.BackColor = Color.FromKnownColor(KnownColor.Window)
    End Sub

    Private Sub btnludatabasesSave_Click(sender As System.Object, e As System.EventArgs) Handles btnludatabasesSave.Click
        Dim iResult As Integer
        Try

            If bModelTypeMod Then
                iResult = SQLHelper.ExecuteScalar(CN, "dbo.s_Insert_Update_Database", _
                                                    IIf(txtdatabasesNamedesc.Text = "", 0, cmbDatabases.SelectedValue(0)), _
                                                    IIf(Trim(txtdatabasesNamedesc.Text) = "", DBNull.Value, LTrim(txtdatabasesNamedesc.Text)), _
                                                    Userid)

                If iResult <> 0 Then
                    MsgBox("Failed to save record change")
                    bModelTypeMod = False
                    Exit Sub
                Else
                    MsgBox("Record Saved")
                End If

                btnludatabasesSave.Visible = False
                btnludatabasesCancel.Visible = False

                dsDatabases = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_database")
                cmbDatabases.DataSource = dsDatabases.Tables(0)
                cmbDatabases.DisplayMember = dsDatabases.Tables(0).Columns("Description").ToString

                txtdatabasesNamedesc.BackColor = Color.FromKnownColor(KnownColor.Window)

                bModelTypeMod = False
            End If

        Catch ex As Exception
            MsgBox("btnludatabasesSave_Click : " + ex.Message)
            'Functions.Sendmail(ex.Message, "btnludatabasesSave_Click " + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub

    Private Sub btnlusoftwareNew_Click(sender As System.Object, e As System.EventArgs) Handles btnluSoftwareNew.Click
        Try
            bInitial = True
            txtSoftwareNamedesc.Clear()
            bInitial = False
            cmbSoftware.SelectedIndex = 0
            txtSoftwareNamedesc.ReadOnly = False

        Catch ex As Exception
            MsgBox("btnlusoftwareNew_Click : New software  : " + ex.Message)
            'Functions.Sendmail(ex.Message, "btnlusoftwareNew_Click " + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub

    Private Sub btnlusoftwareCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnluSoftwareCancel.Click
        cmbsoftware_SelectedIndexChanged(cmbSoftware, EventArgs.Empty)
        txtSoftwareNamedesc.BackColor = Color.FromKnownColor(KnownColor.Window)
    End Sub

    Private Sub btnlusoftwareSave_Click(sender As System.Object, e As System.EventArgs) Handles btnluSoftwareSave.Click
        Dim iResult As Integer
        Try

            If bModelTypeMod Then
                iResult = cmbSoftware.SelectedValue(0)
                iResult = SQLHelper.ExecuteScalar(CN, "dbo.s_Insert_Update_software", _
                                                    IIf(cmbSoftware.SelectedValue(0) = 0, 0, cmbSoftware.SelectedValue(0)), _
                                                    IIf(Trim(txtSoftwareNamedesc.Text) = "", DBNull.Value, LTrim(txtSoftwareNamedesc.Text)), _
                                                    Userid)

                If iResult <> 0 Then
                    MsgBox("Failed to save record change")
                    bModelTypeMod = False
                    Exit Sub
                Else
                    MsgBox("Record Saved")
                End If

                btnluSoftwareSave.Visible = False
                btnluSoftwareCancel.Visible = False

                dsSoftware = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_software")
                cmbSoftware.DataSource = dsSoftware.Tables(0)
                cmbSoftware.DisplayMember = dsSoftware.Tables(0).Columns("Description").ToString

                txtSoftwareNamedesc.BackColor = Color.FromKnownColor(KnownColor.Window)

                bModelTypeMod = False
            End If

        Catch ex As Exception
            MsgBox("btnluSoftwareSave : " + ex.Message)
            'Functions.Sendmail(ex.Message, "btnlusoftwaresSave_Click " + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub

    Private Sub btnluToolsNew_Click(sender As System.Object, e As System.EventArgs) Handles btnluToolsNew.Click
        Try
            bInitial = True
            txtToolsNamedesc.Clear()
            bInitial = False
            cmbTools.SelectedIndex = 0
            txtToolsNamedesc.ReadOnly = False
        Catch ex As Exception
            MsgBox("btnluToolsNew_Click : New Tools  : " + ex.Message)
            'Functions.Sendmail(ex.Message, "btnluToolsNew_Click " + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub

    Private Sub btnluToolsCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnluToolsCancel.Click
        cmbTools_SelectedIndexChanged(cmbTools, EventArgs.Empty)
        txtToolsNamedesc.BackColor = Color.FromKnownColor(KnownColor.Window)
    End Sub

    Private Sub btnluToolsSave_Click(sender As System.Object, e As System.EventArgs) Handles btnluToolsSave.Click
        Dim iResult As Integer
        Try

            If bModelTypeMod Then
                iResult = SQLHelper.ExecuteScalar(CN, "dbo.s_Insert_Update_Tools", _
                                                    IIf(txtToolsNamedesc.Text = "", 0, cmbTools.SelectedValue(0)), _
                                                    IIf(Trim(txtToolsNamedesc.Text) = "", DBNull.Value, LTrim(txtToolsNamedesc.Text)), _
                                                    Userid)

                If iResult <> 0 Then
                    MsgBox("Failed to save record change")
                    bModelTypeMod = False
                    Exit Sub
                Else
                    MsgBox("Record Saved")
                End If

                btnluToolsSave.Visible = False
                btnluToolsCancel.Visible = False

                dsTools = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_Tools")
                cmbTools.DataSource = dsTools.Tables(0)
                cmbTools.DisplayMember = dsTools.Tables(0).Columns("Description").ToString

                txtToolsNamedesc.BackColor = Color.FromKnownColor(KnownColor.Window)

                bModelTypeMod = False
            End If

        Catch ex As Exception
            MsgBox("btnluToolsSave : " + ex.Message)
            'Functions.Sendmail(ex.Message, "btnluToolssSave_Click " + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub

    Private Sub btnluProductNew_Click(sender As System.Object, e As System.EventArgs) Handles btnluProductNew.Click
        Try
            bInitial = True
            txtProductNamedesc.Clear()
            cmbProduct.SelectedIndex = 0
            txtProductNamedesc.ReadOnly = False
            CheckBox29.Enabled = True
            CheckBox29.CheckState = CheckState.Unchecked
            bInitial = False
        Catch ex As Exception
            MsgBox("btnluProductNew_Click : New Product  : " + ex.Message)
            'Functions.Sendmail(ex.Message, "btnluProductNew_Click " + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub

    Private Sub btnNewGroup_Click(sender As System.Object, e As System.EventArgs) Handles btnNewGroup.Click
        Try
            bInitial = True
            txtGroupDesc.Clear()
            cmbGroup.SelectedIndex = 0
            txtGroupDesc.ReadOnly = False
            ckbGroupActive.Enabled = True
            ckbGroupActive.CheckState = CheckState.Unchecked
            bInitial = False
        Catch ex As Exception
            MsgBox("btnNewGroup_Click : New Group  : " + ex.Message)
            'Functions.Sendmail(ex.Message, "btnNewGroup " + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub

    Private Sub btnluProductCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnluProductCancel.Click
        cmbProduct_SelectedIndexChanged(cmbProduct, EventArgs.Empty)
        txtProductNamedesc.BackColor = Color.FromKnownColor(KnownColor.Window)
        CheckBox29.BackColor = Color.Transparent
    End Sub

    Private Sub btnGroupCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnCancelGroup.Click
        cmbGroup_SelectedIndexChanged(cmbGroup, EventArgs.Empty)
        txtGroupDesc.BackColor = Color.FromKnownColor(KnownColor.Window)
        ckbGroupActive.BackColor = Color.Transparent
    End Sub

    Private Sub btnluProductSave_Click(sender As System.Object, e As System.EventArgs) Handles btnluProductSave.Click
        Dim iResult As Integer
        Try

            If bModelTypeMod Then
                bInitial = True
                iResult = SQLHelper.ExecuteScalar(CN, "dbo.s_Insert_Update_Product", _
                                                    IIf(txtProductNamedesc.Text = "", 0, cmbProduct.SelectedValue(0)), _
                                                    IIf(Trim(txtProductNamedesc.Text) = "", DBNull.Value, LTrim(txtProductNamedesc.Text)), _
                                                     IIf(CheckBox29.CheckState = CheckState.Checked, 1, 0), _
                                                    Userid)

                If iResult <> 0 Then
                    MsgBox("Failed to save record change")
                    bModelTypeMod = False
                    Exit Sub
                Else
                    MsgBox("Record Saved")
                End If

                btnluProductSave.Visible = False
                btnluProductCancel.Visible = False

                dsProductList = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_Productlist")
                cmbProduct.DataSource = dsProductList.Tables(0)
                cmbProduct.DisplayMember = dsProductList.Tables(0).Columns("Description").ToString

                dsProjectList = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_ProductList")
                ComboBox2.DataSource = dsProjectList.Tables(0)
                ComboBox2.DisplayMember = dsProjectList.Tables(0).Columns("Description").ToString

                dsFeatureProject = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_ProjectList")
                ComboBox6.DataSource = dsFeatureProject.Tables(0)
                ComboBox6.DisplayMember = dsFeatureProject.Tables(0).Columns("Name").ToString

                txtProductNamedesc.BackColor = Color.FromKnownColor(KnownColor.Window)
                CheckBox29.BackColor = Color.Transparent

                bModelTypeMod = False
                bInitial = False
            End If

        Catch ex As Exception
            MsgBox("btnluProductsSave_Click  : " + ex.Message)
            'Functions.Sendmail(ex.Message, "btnluProductsSave_Click " + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub

    Private Sub btnGroupSave_Click(sender As System.Object, e As System.EventArgs) Handles btnSaveGroup.Click
        Dim iResult As Integer
        Try

            If bModelTypeMod Then
                bInitial = True
                iResult = SQLHelper.ExecuteScalar(CN, "dbo.s_Insert_Update_Group", _
                                                    IIf(txtGroupDesc.Text = "", 0, cmbGroup.SelectedValue(0)), _
                                                    IIf(Trim(txtGroupDesc.Text) = "", DBNull.Value, LTrim(txtGroupDesc.Text)), _
                                                     IIf(ckbGroupActive.CheckState = CheckState.Checked, 1, 0), _
                                                    Userid)

                If iResult <> 0 Then
                    MsgBox("Failed to save record change")
                    bModelTypeMod = False
                    Exit Sub
                Else
                    MsgBox("Record Saved")
                End If

                btnSaveGroup.Visible = False
                btnCancelGroup.Visible = False

                dsGroupList = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_Grouplist")
                cmbGroup.DataSource = dsGroupList.Tables(0)
                cmbGroup.DisplayMember = dsGroupList.Tables(0).Columns("Description").ToString

                txtGroupDesc.BackColor = Color.FromKnownColor(KnownColor.Window)
                ckbGroupActive.BackColor = Color.Transparent

                bModelTypeMod = False
                bInitial = False
            End If

        Catch ex As Exception
            MsgBox("btnluGroupsSave_Click  : " + ex.Message)
            'Functions.Sendmail(ex.Message, "btnluGroupsSave_Click " + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub


    Private Sub btnSaveResource_Click(sender As System.Object, e As System.EventArgs) Handles btnSaveResource.Click
        Dim iResult As Integer, myindex As Integer
        Try

            If bModelTypeMod Then
                bInitial = True
                myindex = cmbResources.SelectedIndex
                iResult = SQLHelper.ExecuteScalar(CN, "dbo.s_Insert_Update_Resource", _
                                                    cmbResources.SelectedValue(0), _
                                                    IIf(txtResource.Text = "", DBNull.Value, LTrim(RTrim(txtResource.Text))), _
                                                    IIf(txtResourceFirstName.Text = "", DBNull.Value, LTrim(RTrim(txtResourceFirstName.Text))), _
                                                    IIf(txtResourceLastName.Text = "", DBNull.Value, LTrim(RTrim(txtResourceLastName.Text))), _
                                                    IIf(txtResourceRole.Text = "", DBNull.Value, LTrim(RTrim(txtResourceRole.Text))), _
                                                    IIf(txtResourceTimezone.Text = "", DBNull.Value, LTrim(RTrim(txtResourceTimezone.Text))), _
                                                    IIf(txtStartTime.Text = "00:00", DBNull.Value, LTrim(RTrim(txtStartTime.Text))), _
                                                    IIf(txtEndTime.Text = "00:00", DBNull.Value, LTrim(RTrim(txtEndTime.Text))), _
                                                    IIf(ckbResourceEngineer.CheckState = CheckState.Checked, 1, 0), _
                                                    IIf(ckbResourceisActive.CheckState = CheckState.Checked, 1, 0), _
                                                    IIf(rtbResourceNote.Text = "", DBNull.Value, LTrim(RTrim(rtbResourceNote.Text))), _
                                                    Userid)

                If iResult <> 0 Then
                    MsgBox("Failed to save record change")
                    bModelTypeMod = False
                    Exit Sub
                Else
                    MsgBox("Record Saved")
                End If

                btnSaveResource.Visible = False
                btnCancelResource.Visible = False

                Reset_color_onResource()

                dsResources = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_ResourceList", IIf(ckbEngineers.CheckState = CheckState.Checked, 1, 0))
                cmbResources.DataSource = dsResources.Tables(0)
                cmbResources.DisplayMember = dsResources.Tables(0).Columns("Resource").ToString

                dsMonthPicker = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_Engineers", 0)
                ComboBox5.DataSource = dsMonthPicker.Tables(0)
                ComboBox5.DisplayMember = dsMonthPicker.Tables(0).Columns("Name").ToString

                'refresh the engineers list just in case we made someone an engineer.
                dsEngineers = SQLHelper.ExecuteDataset(CN, "dbo.s_get_Engineers")

                bModelTypeMod = False


                If myindex = 0 Then
                    cmbResources.SelectedIndex = dsResources.Tables(0).Rows.Count - 1
                Else
                    cmbResources.SelectedIndex = myindex
                End If
                bInitial = False
            End If

        Catch ex As Exception
            MsgBox("btnSaveResource_Click  : " + ex.Message)
            'Functions.Sendmail(ex.Message, "btnSaveResource_Click " + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub

    Private Sub btnNewResource_Click(sender As System.Object, e As System.EventArgs) Handles btnNewResource.Click
        Try
            bInitial = True
            Clear_resource()
            Reset_color_onResource()
            bInitial = False
        Catch ex As Exception
            MsgBox("btnNewResource_Click : New Resource  : " + ex.Message)
            'Functions.Sendmail(ex.Message, "btnNewResource_Click " + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub

    Private Sub btnPMNew_Click(sender As System.Object, e As System.EventArgs) Handles btnPMNew.Click
        Try
            bInitial = True
            txtPMName.Clear()
            txtPMLocation.Clear()
            cmbPM.SelectedIndex = 0
            txtPMLocation.ReadOnly = False
            txtPMName.ReadOnly = False
            ckbPMActive.Enabled = True
            ckbPMActive.CheckState = CheckState.Unchecked
            bInitial = False
        Catch ex As Exception
            MsgBox("btnPMNew_Click : New PM  : " + ex.Message)
            'Functions.Sendmail(ex.Message, "btnPMNew_Click " + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub

    Private Sub btnPMSave_Click(sender As System.Object, e As System.EventArgs) Handles btnPMSave.Click
        Dim iResult As Integer
        Try

            If bModelTypeMod Then
                bInitial = True
                iResult = SQLHelper.ExecuteScalar(CN, "dbo.s_Insert_Update_PM", _
                                                    IIf(txtPMName.Text = "", 0, cmbPM.SelectedValue(0)), _
                                                    IIf(Trim(txtPMName.Text) = "", DBNull.Value, LTrim(txtPMName.Text)), _
                                                    IIf(Trim(txtPMLocation.Text) = "", DBNull.Value, LTrim(txtPMLocation.Text)), _
                                                     IIf(ckbPMActive.CheckState = CheckState.Checked, 1, 0), _
                                                    Userid)

                If iResult <> 0 Then
                    MsgBox("Failed to save record change")
                    bModelTypeMod = False
                    Exit Sub
                Else
                    MsgBox("Record Saved")
                End If

                btnPMSave.Visible = False
                btnPMCancel.Visible = False

                txtPMName.BackColor = Color.FromKnownColor(KnownColor.Window)
                txtPMLocation.BackColor = Color.FromKnownColor(KnownColor.Window)
                ckbPMActive.BackColor = Color.Transparent

                dsPMList = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_PMList")
                cmbPM.DataSource = dsPMList.Tables(0)
                cmbPM.DisplayMember = dsPMList.Tables(0).Columns("Name").ToString

                bModelTypeMod = False
                bInitial = False
                cmbPM.SelectedIndex = 0
            End If

        Catch ex As Exception
            MsgBox("btnluProductsSave_Click  : " + ex.Message)
            'Functions.Sendmail(ex.Message, "btnluProductsSave_Click " + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub

    Private Sub btnSaveProjectMaintenance_Click(sender As System.Object, e As System.EventArgs) Handles btnSaveProjectMaintenance.Click

        Dim iResult As Integer, myindex As Integer
        Try

            If bModelTypeMod Then
                bInitial = True
                Dim sresult As String
                sresult = IIf(TextBox6.Text = "", DBNull.Value, LTrim(RTrim(TextBox6.Text)))
                sresult = IIf(TextBox7.Text = "", DBNull.Value, LTrim(RTrim(TextBox7.Text)))
                sresult = isNumNull(cmbPOList.SelectedValue(0))
                ' sresult = isNumNull(cmbPMList.SelectedValue(0))
                sresult = IIf(TextBox81.Text = "", DBNull.Value, LTrim(RTrim(TextBox81.Text)))
                sresult = IIf(TextBox82.Text = "", DBNull.Value, LTrim(RTrim(TextBox82.Text)))
                sresult = IIf(TextBox83.Text = "", DBNull.Value, LTrim(RTrim(TextBox83.Text)))
                sresult = IIf(TextBox84.Text = "", DBNull.Value, LTrim(RTrim(TextBox84.Text)))
                sresult = IIf(TextBox85.Text = "", DBNull.Value, LTrim(RTrim(TextBox85.Text)))
                sresult = IIf(ckbProjectActive.CheckState = CheckState.Checked, 1, 0)
                sresult = IIf(RichTextBox14.Text = "", DBNull.Value, LTrim(RTrim(RichTextBox14.Text)))
                sresult = IIf(txtPandL.Text = "", DBNull.Value, LTrim(RTrim(txtPandL.Text)))
                sresult = Userid
                myindex = ComboBox2.SelectedIndex
                iResult = SQLHelper.ExecuteScalar(CN, "dbo.s_Insert_Update_Project", _
                                                    ComboBox2.SelectedValue(0), _
                                                    IIf(TextBox6.Text = "", DBNull.Value, LTrim(RTrim(TextBox6.Text))), _
                                                    IIf(TextBox7.Text = "", DBNull.Value, LTrim(RTrim(TextBox7.Text))), _
                                                    isNumNull(cmbPOList.SelectedValue(0)), _
                                                    isNumNull(cmbPMList.SelectedValue(0)), _
                                                    0, _
                                                    0, _
                                                    0, _
                                                    0, _
                                                    0, _
                                                    0, _
                                                    0, _
                                                    0, _
                                                    IIf(TextBox81.Text = "", DBNull.Value, LTrim(RTrim(TextBox81.Text))), _
                                                    IIf(TextBox82.Text = "", DBNull.Value, LTrim(RTrim(TextBox82.Text))), _
                                                    IIf(TextBox83.Text = "", DBNull.Value, LTrim(RTrim(TextBox83.Text))), _
                                                    IIf(TextBox84.Text = "", DBNull.Value, LTrim(RTrim(TextBox84.Text))), _
                                                    IIf(TextBox85.Text = "", DBNull.Value, LTrim(RTrim(TextBox85.Text))), _
                                                    IIf(ckbProjectActive.CheckState = CheckState.Checked, 1, 0), _
                                                    IIf(ckbCapEx.CheckState = CheckState.Checked, 1, 0), _
                                                    IIf(RichTextBox14.Text = "", DBNull.Value, LTrim(RTrim(RichTextBox14.Text))), _
                                                    IIf(txtPandL.Text = "", DBNull.Value, LTrim(RTrim(txtPandL.Text))), _
                                                    Userid)

                If iResult <> 0 Then
                    MsgBox("Failed to save record change")
                    bModelTypeMod = False
                    Exit Sub
                Else
                    MsgBox("Record Saved")
                End If

                btnSaveProjectMaintenance.Visible = False
                btnCancelProjectMaintenance.Visible = False

                Reset_Color_onProject()

                dsProjectList = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_ProjectList")
                ComboBox2.DataSource = dsProjectList.Tables(0)
                ComboBox2.DisplayMember = dsProjectList.Tables(0).Columns("Name").ToString

                dsFeatureProject = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_ProjectList")
                ComboBox6.DataSource = dsFeatureProject.Tables(0)
                ComboBox6.DisplayMember = dsFeatureProject.Tables(0).Columns("Name").ToString

                bModelTypeMod = False
                bInitial = False

                If myindex = 0 Then
                    ComboBox2.SelectedIndex = dsProjectList.Tables(0).Rows.Count - 1
                Else
                    ComboBox2.SelectedIndex = myindex
                End If
            End If

        Catch ex As Exception
            MsgBox("btnSaveProjectMaintenance_Click  : " + ex.Message)
            'Functions.Sendmail(ex.Message, "btnSaveProjectMaintenance_Click " + " : " + Userid, 0, 0, "Project Management")
        End Try

    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        Try
            bInitial = True
            Clear_project()
            ComboBox2.SelectedIndex = 0
            TextBox6.ReadOnly = False
            bInitial = False
        Catch ex As Exception
            MsgBox("Button4_Click : New Product  : " + ex.Message)
            'Functions.Sendmail(ex.Message, "Button4_Click " + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub

    Private Sub btnPMCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnPMCancel.Click
        Try
            bInitial = True
            txtPMName.Clear()
            txtPMLocation.Clear()
            cmbPM.SelectedIndex = 0
            txtPMLocation.ReadOnly = False
            txtPMName.ReadOnly = False
            ckbPMActive.Enabled = True
            ckbPMActive.CheckState = CheckState.Unchecked
            bInitial = False
        Catch ex As Exception
            MsgBox("btnPMCancel_Click : Cancel PM  : " + ex.Message)
            'Functions.Sendmail(ex.Message, "btnPMCancel_Click " + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub

    Private Sub btnPOCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnPOCancel.Click
        Try
            bInitial = True
            txtPOName.Clear()
            txtPOLocation.Clear()
            cmbProjectOwner.SelectedIndex = 0
            txtPOLocation.ReadOnly = False
            txtPOName.ReadOnly = False
            ckbPOActive.Enabled = True
            ckbPOActive.CheckState = CheckState.Unchecked
            bInitial = False
        Catch ex As Exception
            MsgBox("btnPOCancel_Click : Cancel PO  : " + ex.Message)
            'Functions.Sendmail(ex.Message, "btnPOCancel_Click " + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub

    Private Sub btnPOSave_Click(sender As System.Object, e As System.EventArgs) Handles btnPOSave.Click
        Dim iResult As Integer
        Try

            If bModelTypeMod Then
                bInitial = True
                iResult = SQLHelper.ExecuteScalar(CN, "dbo.s_Insert_Update_PO", _
                                                    IIf(txtPOName.Text = "", 0, cmbProjectOwner.SelectedValue(0)), _
                                                    IIf(Trim(txtPOName.Text) = "", DBNull.Value, LTrim(txtPOName.Text)), _
                                                    IIf(Trim(txtPOLocation.Text) = "", DBNull.Value, LTrim(txtPOLocation.Text)), _
                                                     IIf(ckbPOActive.CheckState = CheckState.Checked, 1, 0), _
                                                    Userid)

                If iResult <> 0 Then
                    MsgBox("Failed to save record change")
                    bModelTypeMod = False
                    Exit Sub
                Else
                    MsgBox("Record Saved")
                End If

                btnPOSave.Visible = False
                btnPOCancel.Visible = False

                txtPOName.BackColor = Color.FromKnownColor(KnownColor.Window)
                txtPOLocation.BackColor = Color.FromKnownColor(KnownColor.Window)
                ckbPOActive.BackColor = Color.Transparent

                dsPOList = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_POList")
                cmbProjectOwner.DataSource = dsPOList.Tables(0)
                cmbProjectOwner.DisplayMember = dsPOList.Tables(0).Columns("Name").ToString

                bModelTypeMod = False
                bInitial = False
                cmbProjectOwner.SelectedIndex = 0
            End If

        Catch ex As Exception
            MsgBox("btnPOSave_Click  : " + ex.Message)
            'Functions.Sendmail(ex.Message, "btnPOSave_Click " + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub

    Private Sub btnPONew_Click(sender As System.Object, e As System.EventArgs) Handles btnPONew.Click
        Try
            bInitial = True
            txtPOName.Clear()
            txtPOLocation.Clear()
            cmbProjectOwner.SelectedIndex = 0
            txtPOLocation.ReadOnly = False
            txtPOName.ReadOnly = False
            ckbPOActive.Enabled = True
            ckbPOActive.CheckState = CheckState.Unchecked
            bInitial = False
        Catch ex As Exception
            MsgBox("btnPONew_Click : New PO  : " + ex.Message)
            'Functions.Sendmail(ex.Message, "btnPONew_Click " + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub

    Private Sub btnSaveprojectSprint_Click(sender As System.Object, e As System.EventArgs) Handles btnSaveprojectSprint.Click
        Try



        Catch ex As Exception
            MsgBox("btnSaveprojectSprint_Click  : " + ex.Message)
            'Functions.Sendmail(ex.Message, "btnSaveprojectSprint_Click " + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub


#End Region

#Region "text boxes"
    Private Sub txtdatabasesNamedesc_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtdatabasesNamedesc.TextChanged
        If Not bInitial Then
            btnludatabasesSave.Visible = True
            btnludatabasesCancel.Visible = True
            CType(sender, System.Windows.Forms.TextBox).BackColor = Color.LavenderBlush
            bModelTypeMod = True
        End If
    End Sub

    Private Sub txtsoftwareNamedesc_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtSoftwareNamedesc.TextChanged
        If Not bInitial Then
            btnluSoftwareSave.Visible = True
            btnluSoftwareCancel.Visible = True
            CType(sender, System.Windows.Forms.TextBox).BackColor = Color.LavenderBlush
            bModelTypeMod = True
        End If
    End Sub

    Private Sub txtToolsNamedesc_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtToolsNamedesc.TextChanged
        If Not bInitial Then
            btnluToolsSave.Visible = True
            btnluToolsCancel.Visible = True
            CType(sender, System.Windows.Forms.TextBox).BackColor = Color.LavenderBlush
            bModelTypeMod = True
        End If
    End Sub

    Private Sub txtGroupDesc_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtGroupDesc.TextChanged
        If Not bInitial Then
            btnSaveGroup.Visible = True
            btnCancelGroup.Visible = True
            CType(sender, System.Windows.Forms.TextBox).BackColor = Color.LavenderBlush
            bModelTypeMod = True
        End If
    End Sub

    Private Sub txtProductNamedesc_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtProductNamedesc.TextChanged
        If Not bInitial Then
            btnluProductSave.Visible = True
            btnluProductCancel.Visible = True
            CType(sender, System.Windows.Forms.TextBox).BackColor = Color.LavenderBlush
            bModelTypeMod = True
        End If
    End Sub

    Private Sub txtPM_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtPMName.TextChanged, txtPMLocation.TextChanged
        If Not bInitial Then
            btnPMSave.Visible = True
            btnPMCancel.Visible = True
            CType(sender, System.Windows.Forms.TextBox).BackColor = Color.LavenderBlush
            bModelTypeMod = True
        End If
    End Sub

    Private Sub project_texztbox_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox7.TextChanged, TextBox6.TextChanged, TextBox85.TextChanged, TextBox81.TextChanged, TextBox82.TextChanged, _
                                                                                            TextBox83.TextChanged, TextBox84.TextChanged, txtPandL.TextChanged
        If Not bInitial Then
            btnSaveProjectMaintenance.Visible = True
            btnCancelProjectMaintenance.Visible = True
            CType(sender, System.Windows.Forms.TextBox).BackColor = Color.LavenderBlush
            bModelTypeMod = True
        End If
    End Sub

    Private Sub Project_rtb_TextChanged(sender As System.Object, e As System.EventArgs) Handles RichTextBox14.TextChanged
        If Not bInitial Then
            btnSaveProjectMaintenance.Visible = True
            btnCancelProjectMaintenance.Visible = True
            CType(sender, RichTextBox).BackColor = Color.LavenderBlush
            bModelTypeMod = True
        End If
    End Sub

    Private Sub txtPOName_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtPOName.TextChanged, txtPOLocation.TextChanged
        If Not bInitial Then
            btnPOSave.Visible = True
            btnPOCancel.Visible = True
            CType(sender, System.Windows.Forms.TextBox).BackColor = Color.LavenderBlush
            bModelTypeMod = True
        End If
    End Sub

    Private Sub txtResource_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtResource.TextChanged, txtResourceFirstName.TextChanged, txtResourceLastName.TextChanged, _
                                                                                txtResourceRole.TextChanged, txtResourceTimezone.TextChanged, txtEndTime.TextChanged, txtStartTime.TextChanged
        If Not bInitial Then
            btnSaveResource.Visible = True
            btnCancelResource.Visible = True
            CType(sender, System.Windows.Forms.TextBox).BackColor = Color.LavenderBlush
            bModelTypeMod = True
        End If

    End Sub

    Private Sub rtbResourceNote_TextChanged(sender As System.Object, e As System.EventArgs) Handles rtbResourceNote.TextChanged
        If Not bInitial Then
            btnSaveResource.Visible = True
            btnCancelResource.Visible = True
            CType(sender, System.Windows.Forms.RichTextBox).BackColor = Color.LavenderBlush
            bModelTypeMod = True
        End If
    End Sub

#End Region

#Region "Functions"
    Public Function isnull(ByVal Record As System.Object) As String
        If IsDBNull(Record) Then Return ""
        Return Record
    End Function

    Public Function isNumNull(ByVal Record As System.Object) As Decimal
        If Record.Equals(DBNull.Value) Then
            Return CDec(0)
        Else
            If String.IsNullOrEmpty(CStr(Record)) = True Then
                Return CDec(0)
            Else
                If IsDBNull(Record.ToString) Or Record.ToString Is Nothing Or Record.ToString = "." Then
                    Return CDec(0)
                Else
                    Return CDec(Record)
                End If
            End If
        End If
    End Function

    Public Function isTimeNull(ByVal Record As System.Object) As String
        Try
            If Record.Equals(DBNull.Value) Then
                Return TimeValue("00:00")
            Else
                If String.IsNullOrEmpty(CStr(Record)) = True Then
                    Return TimeValue("00:00")
                Else
                    If IsDBNull(Record.ToString) Or Record.ToString Is Nothing Or Record.ToString = "" Then
                        Return TimeValue("00:00")
                    Else
                        Return TimeValue(Record)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox("Employer Maintenance : isTimeNull :" + ex.Message)
        End Try

    End Function

    Private Sub Clear_project()
        TextBox6.Clear()
        TextBox81.Clear()
        TextBox82.Clear()
        TextBox83.Clear()
        TextBox84.Clear()
        TextBox7.Clear()
        TextBox85.Clear()
        RichTextBox14.Clear()
        txtPandL.Clear()
        ckbProjectActive.CheckState = CheckState.Unchecked
        ckbCapEx.CheckState = CheckState.Unchecked
        cmbPMList.SelectedIndex = 0
        cmbPOList.SelectedIndex = 0
        TextBox72.Clear()
    End Sub

    Private Sub Reset_Color_onProject()
        TextBox6.BackColor = Color.FromKnownColor(KnownColor.Window)
        TextBox81.BackColor = Color.FromKnownColor(KnownColor.Window)
        TextBox82.BackColor = Color.FromKnownColor(KnownColor.Window)
        TextBox83.BackColor = Color.FromKnownColor(KnownColor.Window)
        TextBox84.BackColor = Color.FromKnownColor(KnownColor.Window)
        TextBox7.BackColor = Color.FromKnownColor(KnownColor.Window)
        TextBox85.BackColor = Color.FromKnownColor(KnownColor.Window)
        txtPandL.BackColor = Color.FromKnownColor(KnownColor.Window)
        RichTextBox14.BackColor = Color.FromKnownColor(KnownColor.Window)
        ckbProjectActive.BackColor = Color.Transparent
        ckbCapEx.BackColor = Color.Transparent
    End Sub

    Private Sub Clear_resource()
        cmbResources.SelectedIndex = 0
        txtResource.Clear()
        txtResourceFirstName.Clear()
        txtResourceLastName.Clear()
        txtResourceRole.Clear()
        txtResourceTimezone.Clear()
        txtHrsMTD.Clear()
        txtCapExMTD.Clear()
        txtPctCapEx.Clear()
        txtStartTime.Clear()
        txtEndTime.Clear()
        ckbResourceEngineer.CheckState = CheckState.Unchecked
        ckbResourceisActive.CheckState = CheckState.Unchecked
        rtbResourceNote.Clear()

    End Sub

    Private Sub Reset_color_onResource()
        txtResource.BackColor = Color.FromKnownColor(KnownColor.Window)
        txtResourceFirstName.BackColor = Color.FromKnownColor(KnownColor.Window)
        txtResourceLastName.BackColor = Color.FromKnownColor(KnownColor.Window)
        txtResourceTimezone.BackColor = Color.FromKnownColor(KnownColor.Window)
        txtStartTime.BackColor = Color.FromKnownColor(KnownColor.Window)
        txtEndTime.BackColor = Color.FromKnownColor(KnownColor.Window)
        txtResourceRole.BackColor = Color.FromKnownColor(KnownColor.Window)
        rtbResourceNote.BackColor = Color.FromKnownColor(KnownColor.Window)
        ckbResourceEngineer.BackColor = Color.Transparent
        ckbResourceisActive.BackColor = Color.Transparent
    End Sub

#End Region

#Region "Checkboxes"
    Private Sub CheckBox29_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox29.CheckedChanged
        If Not bInitial Then
            btnluProductSave.Visible = True
            btnluProductCancel.Visible = True
            CType(sender, System.Windows.Forms.CheckBox).BackColor = Color.LavenderBlush
            bModelTypeMod = True
        End If
    End Sub

    Private Sub ckbGroupActive_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles ckbGroupActive.CheckedChanged
        If Not bInitial Then
            btnSaveGroup.Visible = True
            btnCancelGroup.Visible = True
            CType(sender, System.Windows.Forms.CheckBox).BackColor = Color.LavenderBlush
            bModelTypeMod = True
        End If
    End Sub

    Private Sub ckbPMActive_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles ckbPMActive.CheckedChanged
        If Not bInitial Then
            btnPMSave.Visible = True
            btnPMCancel.Visible = True
            CType(sender, System.Windows.Forms.CheckBox).BackColor = Color.LavenderBlush
            bModelTypeMod = True
        End If
    End Sub


    Private Sub CheckBox30_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles ckbProjectActive.CheckedChanged, ckbCapEx.CheckedChanged
        If Not bInitial Then
            btnSaveProjectMaintenance.Visible = True
            btnCancelProjectMaintenance.Visible = True
            CType(sender, System.Windows.Forms.CheckBox).BackColor = Color.LavenderBlush
            bModelTypeMod = True
        End If
    End Sub

    Private Sub ckbPOActive_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles ckbPOActive.CheckedChanged
        If Not bInitial Then
            btnPOSave.Visible = True
            btnPOCancel.Visible = True
            CType(sender, System.Windows.Forms.CheckBox).BackColor = Color.LavenderBlush
            bModelTypeMod = True
        End If
    End Sub

    Private Sub ckbResourceEngineer_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles ckbResourceEngineer.CheckedChanged, ckbResourceisActive.CheckedChanged
        If Not bInitial Then
            btnSaveResource.Visible = True
            btnCancelResource.Visible = True
            CType(sender, System.Windows.Forms.CheckBox).BackColor = Color.LavenderBlush
            bModelTypeMod = True
        End If
    End Sub

    Private Sub ckbEngineers_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles ckbEngineers.CheckedChanged
        Try
            If ckbEngineers.CheckState = CheckState.Checked Then
                dsResources = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_ResourceList", 1)
                cmbResources.DataSource = dsResources.Tables(0)
                cmbResources.DisplayMember = dsResources.Tables(0).Columns("Resource").ToString
            Else
                dsResources = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_ResourceList", 0)
                cmbResources.DataSource = dsResources.Tables(0)
                cmbResources.DisplayMember = dsResources.Tables(0).Columns("Resource").ToString
            End If
        Catch ex As Exception
            'Functions.Sendmail(ex.Message, "ckbEngineers_CheckedChanged", 0, 0, "Employer Maintenance")
            MsgBox("Employer Maintenance : ckbEngineers_CheckedChanged  : " + ex.Message)
        End Try

    End Sub
#End Region

#Region "Datagridview"

    Private Sub dgvActiveTickets_FormatGrid()
        Try
            'set Visual Basic Datagrid Header style to false so we can use our own
            'The key statement required to get the column and row styles to work
            'Visual Header styles must be shut off
            dgvActiveTickets.EnableHeadersVisualStyles = False
            'go and set the styles
            With dgvActiveTickets
                'the following line is necessary for manual column sizing 
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                'let the columns size their heights on their own
                .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
                '*** header settings
                'header backcolor, text color, font bold, font, multiline and alignment
                Dim columnHeaderStyle As New DataGridViewCellStyle
                columnHeaderStyle.BackColor = Color.FromArgb(0, 52, 104)
                columnHeaderStyle.ForeColor = Color.White
                columnHeaderStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                columnHeaderStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                columnHeaderStyle.WrapMode = DataGridViewTriState.True
                'set into place the previously defined header styles
                .ColumnHeadersDefaultCellStyle = columnHeaderStyle
            End With

            'Set DataGridView textbox Column for Product
            Dim colProduct As New DataGridViewTextBoxColumn
            With colProduct
                .DataPropertyName = "Product"
                .HeaderText = "Product"
                .Name = "Product"
                .Visible = True
                .Width = 78
            End With
            dgvActiveTickets.Columns.Add(colProduct)

            'Set DataGridView textbox Column for ParentTicket
            Dim colParentTicket As New DataGridViewTextBoxColumn
            With colParentTicket
                .DataPropertyName = "Parent Ticket"
                .HeaderText = "Parent Ticket"
                .Name = "Parent Ticket"
                .DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                .Width = 90
            End With
            dgvActiveTickets.Columns.Add(colParentTicket)


            'Set DataGridView textbox Column for Ticket
            Dim colTicket As New DataGridViewTextBoxColumn
            With colTicket
                .DataPropertyName = "Ticket"
                .HeaderText = "Ticket"
                .Name = "Ticket"
                .Width = 90
                .DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                ' .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                '.DefaultCellStyle.Format = "##,##0"
            End With
            dgvActiveTickets.Columns.Add(colTicket)

            'Set DataGridView textbox Column for IssueType
            Dim colIssueType As New DataGridViewTextBoxColumn
            With colIssueType
                .DataPropertyName = "IssueType"
                .HeaderText = "Issue Type"
                .Name = "IssueType"
                .Width = 80
                .DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                '.DefaultCellStyle.Format = "##.00"
            End With
            dgvActiveTickets.Columns.Add(colIssueType)

            'Set DataGridView textbox Column for Description
            Dim colDescription As New DataGridViewTextBoxColumn
            With colDescription
                .DataPropertyName = "Description"
                .HeaderText = "Description"
                .Name = "Description"
                .Width = 490
                .DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                '.DefaultCellStyle.Format = "##.00"
            End With
            dgvActiveTickets.Columns.Add(colDescription)

            'Set DataGridView textbox Column for CumulativeHours
            Dim colCumulativeHours As New DataGridViewTextBoxColumn
            With colCumulativeHours
                .DataPropertyName = "Cumulative Hours"
                .HeaderText = "Cumulative Hours"
                .Name = "Cumulative Hours"
                .Width = 80
                .DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                '.DefaultCellStyle.Format = "##.00"
            End With
            dgvActiveTickets.Columns.Add(colCumulativeHours)

        Catch ex As Exception
            Functions.Sendmail(ex.Message, "dgvActiveTickets_FormatGrid", 0, 0, "Employer Maintenance")
            MsgBox("Employer Maintenance : dgvActiveTickets_FormatGrid   :" + ex.Message)
        End Try
    End Sub

    Private Sub dgvActiveTickets_BindData()
        Try
            dgvActiveTickets.DataSource = dsActiveTickets.Tables(0)
        Catch ex As Exception
            Functions.Sendmail(ex.Message, "dgvActiveTickets_BindData", 0, 0, "Employer Maintenance")
            MsgBox("Employer Maintenance : dgvActiveTickets_BindData  : " + ex.Message)
        End Try
    End Sub

    Private Sub dgvMTDHours_FormatGrid()
        Try
            'set Visual Basic Datagrid Header style to false so we can use our own
            'The key statement required to get the column and row styles to work
            'Visual Header styles must be shut off
            dgvMTDHours.EnableHeadersVisualStyles = False
            'go and set the styles
            With dgvMTDHours
                'the following line is necessary for manual column sizing 
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                'let the columns size their heights on their own
                .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
                '*** header settings
                'header backcolor, text color, font bold, font, multiline and alignment
                Dim columnHeaderStyle As New DataGridViewCellStyle
                columnHeaderStyle.BackColor = Color.FromArgb(0, 52, 104)
                columnHeaderStyle.ForeColor = Color.White
                columnHeaderStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                columnHeaderStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                columnHeaderStyle.WrapMode = DataGridViewTriState.True
                'set into place the previously defined header styles
                .ColumnHeadersDefaultCellStyle = columnHeaderStyle
            End With

            'Set DataGridView textbox Column for ResourceID
            Dim colResourceID As New DataGridViewTextBoxColumn
            With colResourceID
                .DataPropertyName = "ResourceID"
                .HeaderText = "ResourceID"
                .Name = "ResourceID"
                .Visible = False
                '.Width = 78
            End With
            dgvMTDHours.Columns.Add(colResourceID)

            'Set DataGridView textbox Column for Resource
            Dim colResource As New DataGridViewTextBoxColumn
            With colResource
                .DataPropertyName = "Resource"
                .HeaderText = "Resource"
                .Name = "Resource"
                .Visible = False
                '.DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                '.Width = 90
            End With
            dgvMTDHours.Columns.Add(colResource)


            'Set DataGridView textbox Column for Product
            Dim colProduct As New DataGridViewTextBoxColumn
            With colProduct
                .DataPropertyName = "Product"
                .HeaderText = "Product"
                .Name = "Product"
                .Width = 90
                .DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
            End With
            dgvMTDHours.Columns.Add(colProduct)

            'Set DataGridView textbox Column for WorkTime
            Dim colWorkTime As New DataGridViewTextBoxColumn
            With colWorkTime
                .DataPropertyName = "WorkTime"
                .HeaderText = "WorkTime"
                .Name = "WorkTime"
                .Width = 90
                .DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                .DefaultCellStyle.Format = "####.00"
            End With
            dgvMTDHours.Columns.Add(colWorkTime)

            'Set DataGridView textbox Column for ProductPercent
            Dim colProductPercent As New DataGridViewTextBoxColumn
            With colProductPercent
                .DataPropertyName = "Product Percent"
                .HeaderText = "Product Percent"
                .Name = "ProductPercent"
                .Width = 90
                .DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                .DefaultCellStyle.Format = "##.00"
            End With
            dgvMTDHours.Columns.Add(colProductPercent)

        Catch ex As Exception
            Functions.Sendmail(ex.Message, "dgvMTDHours_FormatGrid", 0, 0, "Employer Maintenance")
            MsgBox("Employer Maintenance : dgvMTDHours_FormatGrid   :" + ex.Message)
        End Try
    End Sub

    Private Sub dgvMTDHours_BindData()
        Try
            dgvMTDHours.DataSource = dsMTDHours.Tables(0)
        Catch ex As Exception
            Functions.Sendmail(ex.Message, "dgvMTDHours_BindData", 0, 0, "Employer Maintenance")
            MsgBox("Employer Maintenance : dgvMTDHours_BindData  : " + ex.Message)
        End Try
    End Sub

    Private Sub dgvProductsTickets_FormatGrid()
        Try
            'set Visual Basic Datagrid Header style to false so we can use our own
            'The key statement required to get the column and row styles to work
            'Visual Header styles must be shut off
            dgvProductsTickets.EnableHeadersVisualStyles = False
            'go and set the styles
            With dgvProductsTickets
                'the following line is necessary for manual column sizing 
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                'let the columns size their heights on their own
                .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
                '*** header settings
                'header backcolor, text color, font bold, font, multiline and alignment
                Dim columnHeaderStyle As New DataGridViewCellStyle
                columnHeaderStyle.BackColor = Color.FromArgb(0, 52, 104)
                columnHeaderStyle.ForeColor = Color.White
                columnHeaderStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                columnHeaderStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                columnHeaderStyle.WrapMode = DataGridViewTriState.True
                'set into place the previously defined header styles
                .ColumnHeadersDefaultCellStyle = columnHeaderStyle
                .RowHeadersVisible = False
            End With

            'Set DataGridView textbox Column for Resource
            Dim colResource As New DataGridViewTextBoxColumn
            With colResource
                .DataPropertyName = "Resource"
                .HeaderText = "Resource"
                .Name = "Resource"
                .Visible = True
                .DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                .Width = 70
            End With
            dgvProductsTickets.Columns.Add(colResource)


            'Set DataGridView textbox Column for issuekey
            Dim colissuekey As New DataGridViewTextBoxColumn
            With colissuekey
                .DataPropertyName = "issuekey"
                .HeaderText = "Issue Key"
                .Name = "issuekey"
                .Width = 90
                .DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
            End With
            dgvProductsTickets.Columns.Add(colissuekey)

            'Set DataGridView textbox Column for WorkTime
            Dim colWorkTime As New DataGridViewTextBoxColumn
            With colWorkTime
                .DataPropertyName = "WorkTime"
                .HeaderText = "Work Time (H)"
                .Name = "WorkTime"
                .Width = 60
                .DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                .DefaultCellStyle.Format = "####.00"
            End With
            dgvProductsTickets.Columns.Add(colWorkTime)

            'Set DataGridView textbox Column for name
            Dim colname As New DataGridViewTextBoxColumn
            With colname
                .DataPropertyName = "name"
                .HeaderText = "Name"
                .Name = "name"
                .Width = 350
                .DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
            End With
            dgvProductsTickets.Columns.Add(colname)

        Catch ex As Exception
            Functions.Sendmail(ex.Message, "dgvProductsTickets_FormatGrid", 0, 0, "Employer Maintenance")
            MsgBox("Employer Maintenance : dgvProductsTickets_FormatGrid   :" + ex.Message)
        End Try
    End Sub

    Private Sub dgvProductsTickets_BindData()
        Try
            dgvProductsTickets.DataSource = dsProductTickets.Tables(0)
        Catch ex As Exception
            Functions.Sendmail(ex.Message, "dgvProductsTickets_BindData", 0, 0, "Employer Maintenance")
            MsgBox("Employer Maintenance : dgvProductsTickets_BindData  : " + ex.Message)
        End Try
    End Sub

    Private Sub dgvMTDGroupHours_BindData()
        Try
            dgvMTDGroupHours.DataSource = dsGroupMTD.Tables(0)
        Catch ex As Exception
            Functions.Sendmail(ex.Message, "dgvMTDGroupHours_BindData", 0, 0, "Employer Maintenance")
            MsgBox("Employer Maintenance : dgvMTDGroupHours_BindData  : " + ex.Message)
        End Try
    End Sub

    Private Sub dgvProjectMTD_BindData()
        Try
            dgvProjectMTD.DataSource = dsProductMTD.Tables(0)
        Catch ex As Exception
            Functions.Sendmail(ex.Message, "dgvProjectMTD_BindData", 0, 0, "Employer Maintenance")
            MsgBox("Employer Maintenance : dgvProjectMTD_BindData  : " + ex.Message)
        End Try
    End Sub

    Private Sub dgvProjectPrior_BindData()
        Try
            dgvProjectPrior.DataSource = dsProductPrior.Tables(0)
        Catch ex As Exception
            Functions.Sendmail(ex.Message, "dgvProjectPrior_BindData", 0, 0, "Employer Maintenance")
            MsgBox("Employer Maintenance : dgvProjectPrior_BindData  : " + ex.Message)
        End Try
    End Sub

    Private Sub dgvGroupActiveTickets_BindData()
        Try
            dgvGroupActiveTickets.DataSource = dsGroupActiveTickets.Tables(0)
        Catch ex As Exception
            Functions.Sendmail(ex.Message, "dgGroupActiveTickets_BindData", 0, 0, "Employer Maintenance")
            MsgBox("Employer Maintenance : dgvGroupActiveTickets_BindData  : " + ex.Message)
        End Try
    End Sub

    Private Sub dgvTODOTickets_BindData()
        Try
            dgvTODOTickets.DataSource = dsGroupTODOTickets.Tables(0)
        Catch ex As Exception
            Functions.Sendmail(ex.Message, "dgGroupTODOTickets_BindData", 0, 0, "Employer Maintenance")
            MsgBox("Employer Maintenance : dgvTODOTickets_BindData  : " + ex.Message)
        End Try
    End Sub

    Private Sub dgvGroupAssignments_BindData()
        Try
            dgvGroupAssignments.Rows.Clear()
            For i As Integer = 0 To dsGroupAssignments.Tables(0).Rows.Count - 1


                Me.dgvGroupAssignments.Rows.Add(CBool(dsGroupAssignments.Tables(0).Rows(i).Item("isAssigned")), _
                                                dsGroupAssignments.Tables(0).Rows(i).Item(0), _
                                                dsGroupAssignments.Tables(0).Rows(i).Item(1), _
                                                dsGroupAssignments.Tables(0).Rows(i).Item(2), _
                                                dsGroupAssignments.Tables(0).Rows(i).Item(4), _
                                                dsGroupAssignments.Tables(0).Rows(i).Item(5))
                dgvGroupAssignments.Rows(i).Cells("isAssigned").Value = CBool(dsGroupAssignments.Tables(0).Rows(i).Item("isAssigned"))
                If dsGroupAssignments.Tables(0).Rows(i).Item("isused") = 0 Then
                    dgvGroupAssignments.Rows(i).DefaultCellStyle.BackColor = Color.White
                Else
                    dgvGroupAssignments.Rows(i).DefaultCellStyle.BackColor = Color.MintCream
                End If
            Next
            'For Each rw As DataGridViewRow In dgvGroupAssignments.Rows
            '    If rw.Cells("isUsed").Value = 0 Then
            '        dgvGroupAssignments.RowsDefaultCellStyle.BackColor = Color.Transparent
            '    Else
            '        dgvGroupAssignments.DefaultCellStyle.BackColor = Color.Salmon
            '    End If
            'Next
        Catch ex As Exception
            Functions.Sendmail(ex.Message, "dgvGroupAssignments_BindData", 0, 0, "Employer Maintenance")
            MsgBox("Employer Maintenance : dgvGroupAssignments_BindData  : " + ex.Message)
        End Try
    End Sub

    Private Sub dgvMTDGroupHours_FormatGrid()
        Try
            'set Visual Basic Datagrid Header style to false so we can use our own
            'The key statement required to get the column and row styles to work
            'Visual Header styles must be shut off
            dgvMTDGroupHours.EnableHeadersVisualStyles = False
            'go and set the styles
            With dgvMTDGroupHours
                'the following line is necessary for manual column sizing 
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                'let the columns size their heights on their own
                .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
                '*** header settings
                'header backcolor, text color, font bold, font, multiline and alignment
                Dim columnHeaderStyle As New DataGridViewCellStyle
                columnHeaderStyle.BackColor = Color.FromArgb(0, 52, 104)
                columnHeaderStyle.ForeColor = Color.White
                columnHeaderStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                columnHeaderStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                columnHeaderStyle.WrapMode = DataGridViewTriState.True
                'set into place the previously defined header styles
                .ColumnHeadersDefaultCellStyle = columnHeaderStyle
                .RowHeadersVisible = False
            End With

            'Set DataGridView textbox Column for Resource
            Dim colResourceID As New DataGridViewTextBoxColumn
            With colResourceID
                .DataPropertyName = "ResourceID"
                .HeaderText = "ID"
                .Name = "ResourceID"
                .Visible = True
                .DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                .Width = 30
            End With
            dgvMTDGroupHours.Columns.Add(colResourceID)


            'Set DataGridView textbox Column for Name
            Dim colName As New DataGridViewTextBoxColumn
            With colName
                .DataPropertyName = "Name"
                .HeaderText = "Name"
                .Name = "Name"
                .Width = 150
                .DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
            End With
            dgvMTDGroupHours.Columns.Add(colName)

            'Set DataGridView textbox Column for WorkTime
            Dim colWorkTime As New DataGridViewTextBoxColumn
            With colWorkTime
                .DataPropertyName = "WorkTime"
                .HeaderText = "Time"
                .Name = "WorkTime"
                .Width = 60
                .DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                .DefaultCellStyle.Format = "####.00"
            End With
            dgvMTDGroupHours.Columns.Add(colWorkTime)

        Catch ex As Exception
            Functions.Sendmail(ex.Message, "dgvMTDGroupHours_FormatGrid", 0, 0, "Employer Maintenance")
            MsgBox("Employer Maintenance : dgvMTDGroupHours_FormatGrid   :" + ex.Message)
        End Try
    End Sub

    Private Sub dgvPriorGroupHours_BindData()
        Try
            dgvPriorGroupHours.DataSource = dsGroupPrior.Tables(0)
        Catch ex As Exception
            Functions.Sendmail(ex.Message, "dgvPriorGroupHours_BindData", 0, 0, "Employer Maintenance")
            MsgBox("Employer Maintenance : dgvPriorGroupHours_BindData  : " + ex.Message)
        End Try
    End Sub

    Private Sub dgvPriorGroupHours_FormatGrid()
        Try
            'set Visual Basic Datagrid Header style to false so we can use our own
            'The key statement required to get the column and row styles to work
            'Visual Header styles must be shut off
            dgvPriorGroupHours.EnableHeadersVisualStyles = False
            'go and set the styles
            With dgvPriorGroupHours
                'the following line is necessary for manual column sizing 
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                'let the columns size their heights on their own
                .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
                '*** header settings
                'header backcolor, text color, font bold, font, multiline and alignment
                Dim columnHeaderStyle As New DataGridViewCellStyle
                columnHeaderStyle.BackColor = Color.FromArgb(0, 52, 104)
                columnHeaderStyle.ForeColor = Color.White
                columnHeaderStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                columnHeaderStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                columnHeaderStyle.WrapMode = DataGridViewTriState.True
                'set into place the previously defined header styles
                .ColumnHeadersDefaultCellStyle = columnHeaderStyle
                .RowHeadersVisible = False
            End With

            'Set DataGridView textbox Column for Resource
            Dim colResourceID As New DataGridViewTextBoxColumn
            With colResourceID
                .DataPropertyName = "ResourceID"
                .HeaderText = "ID"
                .Name = "ResourceID"
                .Visible = True
                .DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                .Width = 30
            End With
            dgvPriorGroupHours.Columns.Add(colResourceID)


            'Set DataGridView textbox Column for Name
            Dim colName As New DataGridViewTextBoxColumn
            With colName
                .DataPropertyName = "Name"
                .HeaderText = "Name"
                .Name = "Name"
                .Width = 150
                .DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
            End With
            dgvPriorGroupHours.Columns.Add(colName)

            'Set DataGridView textbox Column for WorkTime
            Dim colWorkTime As New DataGridViewTextBoxColumn
            With colWorkTime
                .DataPropertyName = "WorkTime"
                .HeaderText = "Time"
                .Name = "WorkTime"
                .Width = 60
                .DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                .DefaultCellStyle.Format = "####.00"
            End With
            dgvPriorGroupHours.Columns.Add(colWorkTime)

        Catch ex As Exception
            Functions.Sendmail(ex.Message, "dgvPriorGroupHours_FormatGrid", 0, 0, "Employer Maintenance")
            MsgBox("Employer Maintenance : dgvPriorGroupHours_FormatGrid   :" + ex.Message)
        End Try
    End Sub

    Private Sub dgvThisGroupHours_BindData()
        Try
            dgvThisGroupHours.DataSource = dsGroupThis.Tables(0)
        Catch ex As Exception
            Functions.Sendmail(ex.Message, "dgvThisGroupHours_BindData", 0, 0, "Employer Maintenance")
            MsgBox("Employer Maintenance : dgvThisGroupHours_BindData  : " + ex.Message)
        End Try
    End Sub

    Private Sub dgvThisGroupHours_FormatGrid()
        Try
            'set Visual Basic Datagrid Header style to false so we can use our own
            'The key statement required to get the column and row styles to work
            'Visual Header styles must be shut off
            dgvThisGroupHours.EnableHeadersVisualStyles = False
            'go and set the styles
            With dgvThisGroupHours
                'the following line is necessary for manual column sizing 
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                'let the columns size their heights on their own
                .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
                '*** header settings
                'header backcolor, text color, font bold, font, multiline and alignment
                Dim columnHeaderStyle As New DataGridViewCellStyle
                columnHeaderStyle.BackColor = Color.FromArgb(0, 52, 104)
                columnHeaderStyle.ForeColor = Color.White
                columnHeaderStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                columnHeaderStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                columnHeaderStyle.WrapMode = DataGridViewTriState.True
                'set into place the previously defined header styles
                .ColumnHeadersDefaultCellStyle = columnHeaderStyle
                .RowHeadersVisible = False
            End With

            'Set DataGridView textbox Column for Resource
            Dim colResourceID As New DataGridViewTextBoxColumn
            With colResourceID
                .DataPropertyName = "ResourceID"
                .HeaderText = "ID"
                .Name = "ResourceID"
                .Visible = True
                .DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                .Width = 30
            End With
            dgvThisGroupHours.Columns.Add(colResourceID)


            'Set DataGridView textbox Column for Name
            Dim colName As New DataGridViewTextBoxColumn
            With colName
                .DataPropertyName = "Name"
                .HeaderText = "Name"
                .Name = "Name"
                .Width = 150
                .DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
            End With
            dgvThisGroupHours.Columns.Add(colName)

            'Set DataGridView textbox Column for WorkTime
            Dim colWorkTime As New DataGridViewTextBoxColumn
            With colWorkTime
                .DataPropertyName = "WorkTime"
                .HeaderText = "Time"
                .Name = "WorkTime"
                .Width = 60
                .DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                .DefaultCellStyle.Format = "####.00"
            End With
            dgvThisGroupHours.Columns.Add(colWorkTime)

        Catch ex As Exception
            Functions.Sendmail(ex.Message, "dgvThisGroupHours_FormatGrid", 0, 0, "Employer Maintenance")
            MsgBox("Employer Maintenance : dgvThisGroupHours_FormatGrid   :" + ex.Message)
        End Try
    End Sub

    Private Sub dgvProjectMTD_FormatGrid()
        Try
            'set Visual Basic Datagrid Header style to false so we can use our own
            'The key statement required to get the column and row styles to work
            'Visual Header styles must be shut off
            dgvProjectMTD.EnableHeadersVisualStyles = False
            'go and set the styles
            With dgvProjectMTD
                'the following line is necessary for manual column sizing 
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                'let the columns size their heights on their own
                .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
                '*** header settings
                'header backcolor, text color, font bold, font, multiline and alignment
                Dim columnHeaderStyle As New DataGridViewCellStyle
                columnHeaderStyle.BackColor = Color.FromArgb(0, 52, 104)
                columnHeaderStyle.ForeColor = Color.White
                columnHeaderStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                columnHeaderStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                columnHeaderStyle.WrapMode = DataGridViewTriState.True
                'set into place the previously defined header styles
                .ColumnHeadersDefaultCellStyle = columnHeaderStyle
                .RowHeadersVisible = False
            End With



            'Set DataGridView textbox Column for Product
            Dim colProduct As New DataGridViewTextBoxColumn
            With colProduct
                .DataPropertyName = "Product"
                .HeaderText = "Product"
                .Name = "Product"
                .Width = 150
                .DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
            End With
            dgvProjectMTD.Columns.Add(colProduct)

            'Set DataGridView textbox Column for WorkTime
            Dim colWorkTime As New DataGridViewTextBoxColumn
            With colWorkTime
                .DataPropertyName = "WorkTime"
                .HeaderText = "Time"
                .Name = "WorkTime"
                .Width = 60
                .DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                .DefaultCellStyle.Format = "####.00"
            End With
            dgvProjectMTD.Columns.Add(colWorkTime)

        Catch ex As Exception
            Functions.Sendmail(ex.Message, "dgvProjectMTD_FormatGrid", 0, 0, "Employer Maintenance")
            MsgBox("Employer Maintenance : dgvProjectMTD_FormatGrid   :" + ex.Message)
        End Try
    End Sub

    Private Sub dgvProjectPrior_FormatGrid()
        Try
            'set Visual Basic Datagrid Header style to false so we can use our own
            'The key statement required to get the column and row styles to work
            'Visual Header styles must be shut off
            dgvProjectPrior.EnableHeadersVisualStyles = False
            'go and set the styles
            With dgvProjectPrior
                'the following line is necessary for manual column sizing 
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                'let the columns size their heights on their own
                .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
                '*** header settings
                'header backcolor, text color, font bold, font, multiline and alignment
                Dim columnHeaderStyle As New DataGridViewCellStyle
                columnHeaderStyle.BackColor = Color.FromArgb(0, 52, 104)
                columnHeaderStyle.ForeColor = Color.White
                columnHeaderStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                columnHeaderStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                columnHeaderStyle.WrapMode = DataGridViewTriState.True
                'set into place the previously defined header styles
                .ColumnHeadersDefaultCellStyle = columnHeaderStyle
                .RowHeadersVisible = False
            End With



            'Set DataGridView textbox Column for Product
            Dim colProduct As New DataGridViewTextBoxColumn
            With colProduct
                .DataPropertyName = "Product"
                .HeaderText = "Product"
                .Name = "Product"
                .Width = 150
                .DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
            End With
            dgvProjectPrior.Columns.Add(colProduct)

            'Set DataGridView textbox Column for WorkTime
            Dim colWorkTime As New DataGridViewTextBoxColumn
            With colWorkTime
                .DataPropertyName = "WorkTime"
                .HeaderText = "Time"
                .Name = "WorkTime"
                .Width = 60
                .DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                .DefaultCellStyle.Format = "####.00"
            End With
            dgvProjectPrior.Columns.Add(colWorkTime)

        Catch ex As Exception
            Functions.Sendmail(ex.Message, "dgvProjectPrior_FormatGrid", 0, 0, "Employer Maintenance")
            MsgBox("Employer Maintenance : dgvProjectPrior_FormatGrid   :" + ex.Message)
        End Try
    End Sub

    Private Sub dgvGroupActiveTickets_FormatGrid()
        Try
            'set Visual Basic Datagrid Header style to false so we can use our own
            'The key statement required to get the column and row styles to work
            'Visual Header styles must be shut off
            dgvGroupActiveTickets.EnableHeadersVisualStyles = False
            'go and set the styles
            With dgvGroupActiveTickets
                'the following line is necessary for manual column sizing 
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                'let the columns size their heights on their own
                .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
                '*** header settings
                'header backcolor, text color, font bold, font, multiline and alignment
                Dim columnHeaderStyle As New DataGridViewCellStyle
                columnHeaderStyle.BackColor = Color.FromArgb(0, 52, 104)
                columnHeaderStyle.ForeColor = Color.White
                columnHeaderStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                columnHeaderStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                columnHeaderStyle.WrapMode = DataGridViewTriState.True
                'set into place the previously defined header styles
                .ColumnHeadersDefaultCellStyle = columnHeaderStyle
                .RowHeadersVisible = False
            End With

            'Set DataGridView textbox Column for ResourceID
            Dim colResourceID As New DataGridViewTextBoxColumn
            With colResourceID
                .DataPropertyName = "ResourceID"
                .HeaderText = "ResourceID"
                .Name = "ResourceID"
                .Visible = False
                .DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
            End With
            dgvGroupActiveTickets.Columns.Add(colResourceID)

            'Set DataGridView textbox Column for Resource
            Dim colResource As New DataGridViewTextBoxColumn
            With colResource
                .DataPropertyName = "Resource"
                .HeaderText = "Engineer"
                .Name = "Resource"
                .Width = 90
                .DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            End With
            dgvGroupActiveTickets.Columns.Add(colResource)

            'Set DataGridView textbox Column for IssueKey
            Dim colIssueKey As New DataGridViewTextBoxColumn
            With colIssueKey
                .DataPropertyName = "IssueKey"
                .HeaderText = "Issue Key"
                .Name = "IssueKey"
                .Width = 90
                .DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            End With
            dgvGroupActiveTickets.Columns.Add(colIssueKey)

            'Set DataGridView textbox Column for Product
            Dim colProduct As New DataGridViewTextBoxColumn
            With colProduct
                .DataPropertyName = "Product"
                .HeaderText = "Product"
                .Name = "Product"
                .Width = 60
                .DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            End With
            dgvGroupActiveTickets.Columns.Add(colProduct)

            'Set DataGridView textbox Column for Name
            Dim colName As New DataGridViewTextBoxColumn
            With colName
                .DataPropertyName = "Name"
                .HeaderText = "Ticket Name"
                .Name = "Name"
                .Width = 450
                .DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            End With
            dgvGroupActiveTickets.Columns.Add(colName)

            'Set DataGridView textbox Column for WorkTime
            Dim colWorkTime As New DataGridViewTextBoxColumn
            With colWorkTime
                .DataPropertyName = "WorkTime"
                .HeaderText = "Time"
                .Name = "WorkTime"
                .Width = 60
                .DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
                .DefaultCellStyle.Format = "####.00"
            End With
            dgvGroupActiveTickets.Columns.Add(colWorkTime)

            'Set DataGridView textbox Column for isClosed
            Dim colisClosed As New DataGridViewCheckBoxColumn
            With colisClosed
                .DataPropertyName = "isClosed"
                .HeaderText = "is Closed"
                .Name = "isClosed"
                .Width = 60
            End With
            dgvGroupActiveTickets.Columns.Add(colisClosed)


        Catch ex As Exception
            Functions.Sendmail(ex.Message, "dgvGroupActiveTickets_FormatGrid", 0, 0, "Employer Maintenance")
            MsgBox("Employer Maintenance : dgvGroupActiveTickets_FormatGrid   :" + ex.Message)
        End Try
    End Sub

    Private Sub dgvTODOTickets_FormatGrid()
        Try
            'set Visual Basic Datagrid Header style to false so we can use our own
            'The key statement required to get the column and row styles to work
            'Visual Header styles must be shut off
            dgvTODOTickets.EnableHeadersVisualStyles = False
            'go and set the styles
            With dgvTODOTickets
                'the following line is necessary for manual column sizing 
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                'let the columns size their heights on their own
                .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
                '*** header settings
                'header backcolor, text color, font bold, font, multiline and alignment
                Dim columnHeaderStyle As New DataGridViewCellStyle
                columnHeaderStyle.BackColor = Color.FromArgb(0, 52, 104)
                columnHeaderStyle.ForeColor = Color.White
                columnHeaderStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                columnHeaderStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                columnHeaderStyle.WrapMode = DataGridViewTriState.True
                'set into place the previously defined header styles
                .ColumnHeadersDefaultCellStyle = columnHeaderStyle
                .RowHeadersVisible = False
            End With

            'Set DataGridView textbox Column for Name
            Dim colName As New DataGridViewTextBoxColumn
            With colName
                .DataPropertyName = "Name"
                .HeaderText = "Ticket Name"
                .Name = "Name"
                .Width = 450
                .DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            End With
            dgvTODOTickets.Columns.Add(colName)

            'Set DataGridView textbox Column for IssueKey
            Dim colIssueKey As New DataGridViewTextBoxColumn
            With colIssueKey
                .DataPropertyName = "IssueKey"
                .HeaderText = "Issue Key"
                .Name = "IssueKey"
                .Width = 80
                .DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            End With
            dgvTODOTickets.Columns.Add(colIssueKey)

            'Set DataGridView textbox Column for Resource
            Dim colResource As New DataGridViewTextBoxColumn
            With colResource
                .DataPropertyName = "Resource"
                .HeaderText = "Engineer"
                .Name = "Resource"
                .Width = 70
                .DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            End With
            dgvTODOTickets.Columns.Add(colResource)

            'Set DataGridView textbox Column for Product
            Dim colProduct As New DataGridViewTextBoxColumn
            With colProduct
                .DataPropertyName = "Product"
                .HeaderText = "Product"
                .Name = "Product"
                .Width = 70
                .DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            End With
            dgvTODOTickets.Columns.Add(colProduct)


        Catch ex As Exception
            Functions.Sendmail(ex.Message, "dgvTODOTickets_FormatGrid", 0, 0, "Employer Maintenance")
            MsgBox("Employer Maintenance : dgvTODOTickets_FormatGrid   :" + ex.Message)
        End Try
    End Sub

    Private Sub dgvGroupAssignments_FormatGrid()
        Try
            'set Visual Basic Datagrid Header style to false so we can use our own
            'The key statement required to get the column and row styles to work
            'Visual Header styles must be shut off
            dgvGroupAssignments.EnableHeadersVisualStyles = False
            'go and set the styles
            With dgvGroupAssignments
                'the following line is necessary for manual column sizing 
                .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None
                'let the columns size their heights on their own
                .ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
                '*** header settings
                'header backcolor, text color, font bold, font, multiline and alignment
                Dim columnHeaderStyle As New DataGridViewCellStyle
                columnHeaderStyle.BackColor = Color.FromArgb(0, 52, 104)
                columnHeaderStyle.ForeColor = Color.White
                columnHeaderStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                columnHeaderStyle.Alignment = DataGridViewContentAlignment.MiddleCenter
                columnHeaderStyle.WrapMode = DataGridViewTriState.True
                'set into place the previously defined header styles
                .ColumnHeadersDefaultCellStyle = columnHeaderStyle
                .RowHeadersVisible = False
            End With


            'Set DataGridView textbox Column for groupid
            Dim colgroupid As New DataGridViewTextBoxColumn
            With colgroupid
                .DataPropertyName = "groupid"
                .HeaderText = "groupid"
                .Name = "groupid"
                .Visible = False
                '.DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                '.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            End With
            dgvGroupAssignments.Columns.Add(colgroupid)

            'Set DataGridView textbox Column for isAssigned
            Dim colisAssigned As New DataGridViewCheckBoxColumn
            With colisAssigned
                .DataPropertyName = "isAssigned"
                .HeaderText = "Assigned"
                .Name = "isAssigned"
                .Width = 60
                '.DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                '.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            End With
            dgvGroupAssignments.Columns.Add(colisAssigned)


            'Set DataGridView textbox Column for PandL
            Dim colPandL As New DataGridViewTextBoxColumn
            With colPandL
                .DataPropertyName = "PandL"
                .HeaderText = "PandL"
                .Name = "PandL"
                .Width = 80
                .DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            End With
            dgvGroupAssignments.Columns.Add(colPandL)

            'Set DataGridView textbox Column for Name
            Dim colName As New DataGridViewTextBoxColumn
            With colName
                .DataPropertyName = "Name"
                .HeaderText = "Name"
                .Name = "Name"
                .Width = 250
                .DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                .DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            End With
            dgvGroupAssignments.Columns.Add(colName)


            'Set DataGridView textbox Column for isused
            Dim colisused As New DataGridViewTextBoxColumn
            With colisused
                .DataPropertyName = "isused"
                .HeaderText = "isused"
                .Name = "isused"
                .Visible = False
                '.DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                '.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            End With
            dgvGroupAssignments.Columns.Add(colisused)

            'Set DataGridView textbox Column for Projectdimid
            Dim colProjectdimid As New DataGridViewTextBoxColumn
            With colProjectdimid
                .DataPropertyName = "Projectdimid"
                .HeaderText = "Projectdimid"
                .Name = "Projectdimid"
                .Visible = False
                '.DefaultCellStyle.Font = New Font("Arial", 9, FontStyle.Regular)
                '.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft
            End With
            dgvGroupAssignments.Columns.Add(colProjectdimid)


        Catch ex As Exception
            Functions.Sendmail(ex.Message, "dgvGroupAssignments_FormatGrid", 0, 0, "Employer Maintenance")
            MsgBox("Employer Maintenance : dgvGroupAssignments_FormatGrid   :" + ex.Message)
        End Try
    End Sub
#End Region

    Private Sub LinkLabel2_LinkClicked(sender As System.Object, e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        System.Diagnostics.Process.Start("https://confluence.teamdrg.com/display/PT/Product+Roadmaps")

    End Sub

    Private Sub LinkLabel3_LinkClicked(sender As System.Object, e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        System.Diagnostics.Process.Start("https://docs.google.com/spreadsheets/d/1UrdzE5yvozVc2qzFUTJtGk-fHoyLHZOqjujU7EAxp7c/edit#gid=1381192490")
    End Sub

    'Private Sub Form1_Shown(ByVal sender As Object, _
    '                        ByVal e As System.EventArgs) _
    '                    Handles Me.Shown
    '    'set the DateTimePicker to display only time
    '    dtpStart.Format = DateTimePickerFormat.Custom
    '    dtpStart.CustomFormat = myFormat
    '    dtpStart.ShowUpDown = True

    'End Sub

    'Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, _
    '                                         ByVal e As System.EventArgs) _
    '                                     Handles dtpStart.ValueChanged
    '    'this shows how to extract time from the DateTimePicker
    '    TextBox1.Text = dtpStart.Value.ToString(myFormat)

    '    Dim ts As TimeSpan = dtpStart.Value.TimeOfDay 'the time of day only
    '    If Not bInitial Then
    '        btnSaveResource.Visible = True
    '        btnCancelResource.Visible = True
    '        CType(sender, DateTimePicker).BackColor = Color.LavenderBlush
    '        bModelTypeMod = True
    '    End If
    'End Sub

    Private Sub cmbGroupLead_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbGroupLead.SelectedIndexChanged, cmbGroupDB1.SelectedIndexChanged, cmbGroupDB2.SelectedIndexChanged, _
                                                                                                            cmbGroupDB3.SelectedIndexChanged
        Try
            If Not bInitial Then
                bInitial = True
                btnGroupAssingmentSave.Visible = True
                bModelTypeMod = True
                bInitial = False
            End If

        Catch ex As Exception
            'Functions.Sendmail(ex.Message, "cmbsoftware_SelectedIndexChanged" + " : " + Userid, 0, 0, "Project Management")
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub btnGroupAssingmentSave_Click(sender As System.Object, e As System.EventArgs) Handles btnGroupAssingmentSave.Click
        Dim iResult As Integer, myindex As Integer
        Try

            If bModelTypeMod Then
                bInitial = True
                myindex = cmbGroupLead.SelectedValue
                myindex = cmbGroupDB1.SelectedValue
                myindex = cmbGroupDB2.SelectedValue
                myindex = cmbGroupDB3.SelectedValue
                myindex = cmbGroups.SelectedValue(0)
                iResult = SQLHelper.ExecuteScalar(CN, "dbo.s_Insert_Update_GroupAssignment", _
                                                    isNumNull(cmbGroups.SelectedValue(0)), _
                                                    isNumNull(cmbGroupLead.SelectedValue), _
                                                    isNumNull(cmbGroupDB1.SelectedValue), _
                                                    isNumNull(cmbGroupDB2.SelectedValue), _
                                                    isNumNull(cmbGroupDB3.SelectedValue), _
                                                    Userid)

                If iResult <> 0 Then
                    MsgBox("Failed to save record change")
                    bModelTypeMod = False
                    Exit Sub
                Else
                    MsgBox("Record Saved")
                End If

                btnGroupAssingmentSave.Visible = False

                Reset_color_onResource()

                'dsResources = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_ResourceList", IIf(ckbEngineers.CheckState = CheckState.Checked, 1, 0))
                'cmbResources.DataSource = dsResources.Tables(0)
                'cmbResources.DisplayMember = dsResources.Tables(0).Columns("Resource").ToString

                bModelTypeMod = False


                'If myindex = 0 Then
                '    cmbResources.SelectedIndex = dsResources.Tables(0).Rows.Count - 1
                'Else
                '    cmbResources.SelectedIndex = myindex
                'End If
                bInitial = False
            End If

        Catch ex As Exception
            MsgBox("btnGroupAssingmentSave_Click  : " + ex.Message)
            'Functions.Sendmail(ex.Message, "btnGroupAssingmentSave_Click " + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub

    Public Sub OpenExcelDemo(ByVal FileName As String, ByVal SheetName As String)
        If IO.File.Exists(FileName) Then
            Dim Proceed As Boolean = False
            Dim xlApp As Excel.Application = Nothing
            Dim xlWorkBooks As Excel.Workbooks = Nothing
            Dim xlWorkBook As Excel.Workbook = Nothing
            Dim xlWorkSheet As Excel.Worksheet = Nothing
            Dim xlWorkSheets As Excel.Sheets = Nothing
            Dim xlCells As Excel.Range = Nothing

            xlApp = New Excel.Application
            xlApp.DisplayAlerts = False
            xlWorkBooks = xlApp.Workbooks
            xlWorkBook = xlWorkBooks.Open(FileName)
            xlWorkBook.RefreshAll()
            xlApp.Visible = True
            xlWorkSheets = xlWorkBook.Sheets
            For x As Integer = 1 To xlWorkSheets.Count
                xlWorkSheet = CType(xlWorkSheets(x), Excel.Worksheet)
                If xlWorkSheet.Name = SheetName Then
                    Console.WriteLine(SheetName)
                    Proceed = True
                    Exit For
                End If
                Runtime.InteropServices.Marshal.FinalReleaseComObject(xlWorkSheet)
                xlWorkSheet = Nothing
            Next
            If Proceed Then
                xlWorkSheet.Activate()
                MessageBox.Show("File is open, if you close Excel just opened outside of this program we will crash-n-burn.")
            Else
                MessageBox.Show(SheetName & " not found.")
            End If
            xlWorkBook.Close()

            xlApp.UserControl = True
            xlApp.Quit()
            ReleaseComObject(xlCells)
            ReleaseComObject(xlWorkSheets)
            ReleaseComObject(xlWorkSheet)
            ReleaseComObject(xlWorkBook)
            ReleaseComObject(xlWorkBooks)
            ReleaseComObject(xlApp)
            'KillProcessByName("excel.exe")
            GC.Collect()
        Else
            MessageBox.Show("'" & FileName & "' not located. Try one of the write examples first.")
        End If
    End Sub

    Public Sub ReleaseComObject(ByVal obj As Object)
        Try
            If obj IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
                obj = Nothing
            End If

        Catch ex As Exception
            obj = Nothing
        End Try
    End Sub

    Public Sub KillProcessByName(ByVal psProcessName As String)

        '* Iterate through all running processes to locate the specified process and kill it.
        For Each pProcess As Process In Process.GetProcessesByName(psProcessName)
            '* Kill the process
            If Not pProcess.HasExited Then Call pProcess.Kill()
            '* Wait while the process completes it's exit.
            Do While Not pProcess.HasExited
                Call System.Windows.Forms.Application.DoEvents()
            Loop
        Next 'pProcess

    End Sub

    Private Sub LinkLabel5_LinkClicked(sender As System.Object, e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel5.LinkClicked
        System.Diagnostics.Process.Start("http://myreports/DRG/Pages/Report.aspx?ItemPath=%2fStaging%2fJira%2fTickets+with+no+Project&SelectedSubTabId=ReportDataSourcePropertiesTab")
    End Sub

    Private Sub LinkLabel4_LinkClicked(sender As System.Object, e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel4.LinkClicked
        System.Diagnostics.Process.Start("http://myreports/DRG/Pages/Report.aspx?ItemPath=%2fStaging%2fJira%2fUn-Assigned+Projects")
    End Sub

    Private Sub dgvGroupAssignments_Cellvaluechanged(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvGroupAssignments.CellValueChanged
        Try
            If Not bInitial Then
                Dim mystr As String = dgvGroupAssignments.Rows(e.RowIndex).Cells("ProjectDimID").Value.ToString
                mystr = dgvGroupAssignments.Rows(e.RowIndex).Cells(6).Value.ToString
                GlobalLibrary.SqlHelper.ExecuteNonQuery(CN, "dbo.s_Insert_Update_GroupAssignments", _
                                                        isNumNull(dgvGroupAssignments.Rows(e.RowIndex).Cells("groupid")), _
                                                        isNumNull(dgvGroupAssignments.Rows(e.RowIndex).Cells("ProjectDimID").Value), _
                                                        CBool(dgvGroupAssignments.Rows(e.RowIndex).Cells("isAssigned").Value), _
                                                        Userid)
            End If
            'GlobalLibrary.SqlHelper.ExecuteNonQuery(CN, "dbo.s_Insert_Update_GroupAssignments", _
            '                            dgvGroupAssignments.Rows(e.RowIndex).Cells("groupid").ToString, _
            '                            dgvGroupAssignments.Rows(e.RowIndex).Cells("ProjectDimID").Value.ToString, _
            '                            CBool(dgvGroupAssignments.Rows(e.RowIndex).Cells("isAssigned").Value), _
            '                            Userid)
        Catch ex As Exception
            MsgBox("dgvGroupAssignments_Cellvaluechanged" + " : " + ex.Message)
            'Functions.Sendmail(ex.Message, "dgvGroupAssignments_Cellvaluechanged" + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub

#Region "Feature Dim maint"

    Private Sub btnSaveFeature_Click(sender As System.Object, e As System.EventArgs)
        Dim iResult As Integer
        Try

            If bModelTypeMod Then
                bInitial = True
                iResult = SQLHelper.ExecuteScalar(CN, "dbo.s_Insert_Update_Product", _
                                                    IIf(TextBox20.Text = "", 0, cmbfeature.SelectedValue(0)), _
                                                    IIf(Trim(TextBox20.Text) = "", DBNull.Value, LTrim(TextBox20.Text)), _
                                                    IIf(txtFeaturePriority.Text = "", 99, txtFeaturePriority.Text), _
                                                     IIf(CheckBox1.CheckState = CheckState.Checked, 1, 0), _
                                                    Userid)

                If iResult <> 0 Then
                    MsgBox("Failed to save record change")
                    bModelTypeMod = False
                    Exit Sub
                Else
                    MsgBox("Record Saved")
                End If

                btnSaveFeature.Visible = False
                btnCancelFeature.Visible = False

                dsFeatureList = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_FeatureList", 0)
                cmbfeature.DataSource = dsFeatureList.Tables(0)
                cmbfeature.DisplayMember = dsFeatureList.Tables(0).Columns("Description").ToString

                'dsFeature = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_FeatureList", 1)
                'cmbfeatureMgmt.DataSource = dsFeature.Tables(0)
                'cmbfeatureMgmt.DisplayMember = dsFeature.Tables(0).Columns("Description").ToString

                TextBox20.BackColor = Color.FromKnownColor(KnownColor.Window)
                txtFeaturePriority.BackColor = Color.FromKnownColor(KnownColor.Window)
                CheckBox1.BackColor = Color.Transparent

                bModelTypeMod = False
                bInitial = False
            End If

        Catch ex As Exception
            MsgBox("btnluProductsSave_Click  : " + ex.Message)
            'Functions.Sendmail(ex.Message, "btnluProductsSave_Click " + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub

    Private Sub cmbfeature_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cmbfeature.SelectedIndexChanged
        Dim myindex As Integer
        Try
            If Not bInitial Then
                bInitial = True
                myindex = cmbfeature.SelectedIndex
                dsfeatureitem = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_FeatureDim", cmbfeature.SelectedValue(0))

                If dsfeatureitem.Tables(0).Rows.Count > 0 Then

                    TextBox20.Text = isnull(dsfeatureitem.Tables(0).Rows(0).Item("Description"))                      'ComboBox6
                    txtFeaturePriority.Text = isnull(dsfeatureitem.Tables(0).Rows(0).Item("Priority"))
                    Functions.whyareyousodumb(CheckBox1, dsfeatureitem.Tables(0).Rows(0).Item("isActive"))
                    TextBox8.Text = isnull(dsfeatureitem.Tables(0).Rows(0).Item("CapEx"))
                    TextBox28.Text = isnull(dsfeatureitem.Tables(0).Rows(0).Item("DK Ticket"))
                    TextBox36.Text = isnull(dsfeatureitem.Tables(0).Rows(0).Item("PandL"))
                    TextBox20.ReadOnly = False
                    txtFeaturePriority.ReadOnly = False
                    TextBox8.ReadOnly = False
                    btnSaveFeature.Visible = False
                    btnCancelFeature.Visible = False
                    CheckBox1.Enabled = True
                    cmbfeature.SelectedIndex = myindex
                    TextBox20.BackColor = Color.FromKnownColor(KnownColor.Window)
                    TextBox8.BackColor = Color.FromKnownColor(KnownColor.Window)
                    txtFeaturePriority.BackColor = Color.FromKnownColor(KnownColor.Window)
                    CheckBox1.BackColor = Color.Transparent
                    TextBox28.BackColor = Color.FromKnownColor(KnownColor.Window)
                    TextBox36.BackColor = Color.FromKnownColor(KnownColor.Window)
                    ComboBox6.BackColor = Color.FromKnownColor(KnownColor.Window)

                    If isNumNull(dsfeatureitem.Tables(0).Rows(0).Item("ProjectDimID")) = 0 Then
                        ComboBox6.SelectedIndex = 0
                    Else
                        ComboBox6.SelectedIndex = ComboBox6.FindStringExact(dsfeatureitem.Tables(0).Rows(0).Item("ProjectName"))
                    End If

                    ' Re-set the fact management box
                    ComboBox4.SelectedIndex = 0
                    ComboBox5.SelectedIndex = 0
                    TextBox14.Clear()

                    dsfeaturebreakout = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_feature_breakout", cmbfeature.SelectedValue(0))
                    If dsfeaturebreakout.Tables.Count > 0 Then
                        If dsfeaturebreakout.Tables(0).Rows.Count > 0 Then
                            DataGridView2.DataSource = dsfeaturebreakout.Tables(0)
                        Else
                            DataGridView2.DataSource = Nothing
                        End If
                    Else
                        DataGridView2.DataSource = Nothing
                    End If

                    dsfeaturetime = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_rpt_feature_time")
                    If dsfeaturetime.Tables.Count > 0 Then
                        If dsfeaturetime.Tables(0).Rows.Count > 0 Then
                            DataGridView3.DataSource = dsfeaturetime.Tables(0)
                        Else
                            DataGridView3.DataSource = Nothing
                        End If
                    Else
                        DataGridView3.DataSource = Nothing
                    End If

                    dsmissingtickets = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_rpt_missing_tickets")
                    If dsmissingtickets.Tables.Count > 0 Then
                        If dsmissingtickets.Tables(0).Rows.Count > 0 Then
                            DataGridView5.DataSource = dsmissingtickets.Tables(0)
                        Else
                            DataGridView5.DataSource = Nothing
                        End If
                    Else
                        DataGridView5.DataSource = Nothing
                    End If


                Else
                    TextBox20.Clear()
                    TextBox8.Clear()
                    txtFeaturePriority.Clear()
                    CheckBox1.Checked = False
                    ComboBox6.SelectedIndex = 0
                    TextBox28.Clear()
                    TextBox36.Clear()
                    DataGridView2.DataSource = Nothing
                    DataGridView3.DataSource = Nothing
                    DataGridView5.DataSource = Nothing
                End If
                bInitial = False
            End If




        Catch ex As Exception
            MsgBox("cmbfeature_SelectedIndexChanged" + " : " + ex.Message)
            'Functions.Sendmail(ex.Message, "cmbfeature_SelectedIndexChanged" + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub

    Private Sub btnNewFeature_Click(sender As System.Object, e As System.EventArgs)
        Try
            bInitial = True
            TextBox20.Clear()
            TextBox8.Clear()
            txtFeaturePriority.Clear()
            cmbfeature.SelectedIndex = 0
            TextBox20.ReadOnly = False
            TextBox8.ReadOnly = False
            txtFeaturePriority.ReadOnly = False
            CheckBox1.Enabled = True
            CheckBox1.CheckState = CheckState.Unchecked
            bInitial = False
        Catch ex As Exception
            MsgBox("btnNewFeature_Click : New Product  : " + ex.Message)
            'Functions.Sendmail(ex.Message, "btnNewFeature_Click " + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub

    Private Sub btnCancelFeature_Click(sender As System.Object, e As System.EventArgs)
        cmbfeature_SelectedIndexChanged(cmbfeature, EventArgs.Empty)
        TextBox20.BackColor = Color.FromKnownColor(KnownColor.Window)
        TextBox8.BackColor = Color.FromKnownColor(KnownColor.Window)
        txtFeaturePriority.BackColor = Color.FromKnownColor(KnownColor.Window)
        CheckBox1.BackColor = Color.Transparent
    End Sub

    Private Sub TextBox20_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox20.TextChanged, TextBox8.TextChanged
        If Not bInitial Then
            btnSaveFeature.Visible = True
            btnCancelFeature.Visible = True
            CType(sender, System.Windows.Forms.TextBox).BackColor = Color.LavenderBlush
            bModelTypeMod = True
        End If
    End Sub

    Private Sub btnFeatureCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnCancelFeature.Click
        Try
            bInitial = True
            TextBox20.Clear()
            TextBox8.Clear()
            txtFeaturePriority.Clear()
            cmbfeature.SelectedIndex = 0
            TextBox20.ReadOnly = False
            TextBox8.ReadOnly = False
            txtFeaturePriority.ReadOnly = False
            CheckBox1.Enabled = True
            CheckBox1.CheckState = CheckState.Unchecked
            bInitial = False
        Catch ex As Exception
            MsgBox("btnfeatureCancel_Click : Cancel PO  : " + ex.Message)
            'Functions.Sendmail(ex.Message, "btnfeatureCancel_Click " + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub

    Private Sub btnFeatureSave_Click(sender As System.Object, e As System.EventArgs) Handles btnSaveFeature.Click
        Dim iResult As Integer
        Try

            If bModelTypeMod Then
                bInitial = True
                Dim ir As Integer = isNumNull(ComboBox6.SelectedValue(0))
                iResult = SQLHelper.ExecuteScalar(CN, "dbo.s_Insert_Update_Feature", _
                                                    IIf(TextBox20.Text = "", 0, cmbfeature.SelectedValue(0)), _
                                                    IIf(Trim(TextBox20.Text) = "", DBNull.Value, LTrim(TextBox20.Text)), _
                                                    IIf(CheckBox1.CheckState = CheckState.Checked, 1, 0), _
                                                    IIf(Trim(txtFeaturePriority.Text) = "", DBNull.Value, LTrim(txtFeaturePriority.Text)), _
                                                    IIf(Trim(TextBox8.Text) = "", DBNull.Value, LTrim(TextBox8.Text)), _
                                                    isNumNull(ComboBox6.SelectedValue(0)), _
                                                    IIf(Trim(TextBox28.Text) = "", DBNull.Value, LTrim(TextBox28.Text)), _
                                                    IIf(Trim(TextBox36.Text) = "", DBNull.Value, LTrim(TextBox36.Text)), _
                                                    Userid)

                If iResult <> 0 Then
                    MsgBox("Failed to save record change")
                    bModelTypeMod = False
                    Exit Sub
                Else
                    MsgBox("Record Saved")
                End If

                btnSaveFeature.Visible = False
                btnCancelFeature.Visible = False

                TextBox20.BackColor = Color.FromKnownColor(KnownColor.Window)
                TextBox8.BackColor = Color.FromKnownColor(KnownColor.Window)
                txtFeaturePriority.BackColor = Color.FromKnownColor(KnownColor.Window)
                CheckBox1.BackColor = Color.Transparent
                TextBox28.BackColor = Color.FromKnownColor(KnownColor.Window)
                TextBox36.BackColor = Color.FromKnownColor(KnownColor.Window)
                ComboBox6.BackColor = Color.FromKnownColor(KnownColor.Window)

                dsFeatureList = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_FeatureList", 0)
                cmbfeature.DataSource = dsFeatureList.Tables(0)
                cmbfeature.DisplayMember = dsFeatureList.Tables(0).Columns("Description").ToString

                bModelTypeMod = False

                TextBox20.Clear()
                TextBox8.Clear()
                TextBox28.Clear()
                TextBox36.Clear()
                txtFeaturePriority.Clear()
                cmbfeature.SelectedIndex = 0
                ComboBox6.SelectedIndex = 0
                TextBox20.ReadOnly = False
                TextBox8.ReadOnly = False
                txtFeaturePriority.ReadOnly = False
                CheckBox1.Enabled = True
                CheckBox1.CheckState = CheckState.Unchecked
                bInitial = False

                dsfeaturebreakout = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_feature_breakout", cmbfeature.SelectedValue(0))
                If dsfeaturebreakout.Tables.Count > 0 Then
                    If dsfeaturebreakout.Tables(0).Rows.Count > 0 Then
                        DataGridView2.DataSource = dsfeaturebreakout.Tables(0)
                    Else
                        DataGridView2.DataSource = Nothing
                    End If
                End If

                dsfeaturetime = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_rpt_feature_time")
                If dsfeaturetime.Tables.Count > 0 Then
                    If dsfeaturetime.Tables(0).Rows.Count > 0 Then
                        DataGridView3.DataSource = dsfeaturetime.Tables(0)
                    Else
                        DataGridView3.DataSource = Nothing
                    End If
                End If

                dsmissingtickets = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_rpt_missing_tickets")
                If dsmissingtickets.Tables.Count > 0 Then
                    If dsmissingtickets.Tables(0).Rows.Count > 0 Then
                        DataGridView5.DataSource = dsmissingtickets.Tables(0)
                    Else
                        DataGridView5.DataSource = Nothing
                    End If
                End If

            End If

        Catch ex As Exception
            MsgBox("btnFeatureSave_Click  : " + ex.Message)
            'Functions.Sendmail(ex.Message, "btnFeatureSave_Click " + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub

    Private Sub btnFeatureNew_Click(sender As System.Object, e As System.EventArgs) Handles btnNewFeature.Click
        Try
            bInitial = True
            TextBox20.Clear()
            TextBox8.Clear()
            txtFeaturePriority.Clear()
            cmbfeature.SelectedIndex = 0
            TextBox20.ReadOnly = False
            TextBox8.ReadOnly = False
            CheckBox1.Enabled = True
            txtFeaturePriority.Enabled = True
            CheckBox1.CheckState = CheckState.Unchecked
            TextBox28.Clear()
            TextBox36.Clear()
            TextBox28.ReadOnly = False
            TextBox36.ReadOnly = False
            ComboBox6.SelectedIndex = 0
            DataGridView2.DataSource = Nothing
            DataGridView3.DataSource = Nothing
            DataGridView5.DataSource = Nothing
            bInitial = False
        Catch ex As Exception
            MsgBox("btnFeatureNew_Click : New Feature  : " + ex.Message)
            'Functions.Sendmail(ex.Message, "btnFeatureNew_Click " + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub

    Private Sub ckbFeatureActive_CheckedChanged(sender As System.Object, e As System.EventArgs)
        If Not bInitial Then
            btnSaveFeature.Visible = True
            btnCancelFeature.Visible = True
            CType(sender, System.Windows.Forms.CheckBox).BackColor = Color.LavenderBlush
            bModelTypeMod = True
        End If
    End Sub

    Private Sub txtfeatureName_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox20.TextChanged, txtFeaturePriority.TextChanged, TextBox8.TextChanged, TextBox28.TextChanged, TextBox36.TextChanged
        If Not bInitial Then
            btnSaveFeature.Visible = True
            btnCancelFeature.Visible = True
            CType(sender, System.Windows.Forms.TextBox).BackColor = Color.LavenderBlush
            bModelTypeMod = True
        End If
    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox6.SelectedIndexChanged
        If Not bInitial Then
            btnSaveFeature.Visible = True
            btnCancelFeature.Visible = True
            CType(sender, System.Windows.Forms.ComboBox).BackColor = Color.LavenderBlush
            bModelTypeMod = True
        End If
    End Sub

    Private Sub btnSaveFeaturefact_Click(sender As System.Object, e As System.EventArgs) Handles btnSaveFeaturefact.Click
        Dim iResult As Integer
        Try

            If bModelTypeMod Then
                bInitial = True
                iResult = SQLHelper.ExecuteScalar(CN, "dbo.s_Insert_Update_FeatureFact", _
                                                    isNumNull(cmbfeature.SelectedValue(0)), _
                                                    isNumNull(ComboBox4.SelectedValue(0)), _
                                                    isNumNull(ComboBox5.SelectedValue(0)), _
                                                    IIf(Trim(TextBox14.Text) = "", DBNull.Value, LTrim(TextBox14.Text)), _
                                                    Userid)

                If iResult <> 0 Then
                    MsgBox("Failed to save record change")
                    bModelTypeMod = False
                    Exit Sub
                Else
                    MsgBox("Record Saved")
                End If

                ComboBox4.BackColor = Color.FromKnownColor(KnownColor.Window)
                ComboBox5.BackColor = Color.FromKnownColor(KnownColor.Window)
                TextBox14.BackColor = Color.FromKnownColor(KnownColor.Window)


                'Update the grid with the new data
                'This is really tricky because month columns are dynamic
                'also need to verify that a resource is not over 100% for a given month
                'also need to track resource utilization

                dsfeaturebreakout = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_get_feature_breakout", cmbfeature.SelectedValue(0))
                If dsfeaturebreakout.Tables.Count > 0 Then
                    If dsfeaturebreakout.Tables(0).Rows.Count > 0 Then
                        DataGridView2.DataSource = dsfeaturebreakout.Tables(0)
                    Else
                        DataGridView2.DataSource = Nothing
                    End If
                End If

                dsfeaturetime = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_rpt_feature_time")
                If dsfeaturetime.Tables.Count > 0 Then
                    If dsfeaturetime.Tables(0).Rows.Count > 0 Then
                        DataGridView3.DataSource = dsfeaturetime.Tables(0)
                    Else
                        DataGridView3.DataSource = Nothing
                    End If
                End If

                dsmissingtickets = GlobalLibrary.SqlHelper.ExecuteDataset(CN, "dbo.s_rpt_missing_tickets")
                If dsmissingtickets.Tables.Count > 0 Then
                    If dsmissingtickets.Tables(0).Rows.Count > 0 Then
                        DataGridView5.DataSource = dsmissingtickets.Tables(0)
                    Else
                        DataGridView5.DataSource = Nothing
                    End If
                End If

                bModelTypeMod = False

                bInitial = False

            End If

        Catch ex As Exception
            MsgBox("btnFeatureSave_Click  : " + ex.Message)
            'Functions.Sendmail(ex.Message, "btnFeatureSave_Click " + " : " + Userid, 0, 0, "Project Management")
        End Try
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox4.SelectedIndexChanged
        If Not bInitial Then
            bModelTypeMod = True
        End If
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBox5.SelectedIndexChanged
        If Not bInitial Then
            bModelTypeMod = True
        End If
    End Sub

    Private Sub TextBox14_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox14.TextChanged
        If Not bInitial Then
            CType(sender, System.Windows.Forms.TextBox).BackColor = Color.LavenderBlush
            bModelTypeMod = True
        End If
    End Sub

#End Region

#Region "Feature Management"

    'Private Sub cmbfeatureMgmt_SelectedIndexChanged(sender As System.Object, e As System.EventArgs)

    '    Try
    '        If Not bInitial Then
    '            bInitial = True

    '            dsFeatureManagement = SQLHelper.ExecuteDataset(CN, "dbo.s_get_Feature", cmbfeatureMgmt.SelectedValue(0))
    '            Label26.Text = cmbfeatureMgmt.SelectedValue(0)
    '            If dsFeatureManagement.Tables(0).Rows.Count > 0 Then
    '                TextBox8.Text = isnull(dsFeatureManagement.Tables(0).Rows(0).Item("Capex"))

    '                If dsFeatureManagement.Tables(0).Rows(0).Item("projectdimid") = 0 Then
    '                    ComboBox3.SelectedIndex = 0
    '                    'TextBox72.Clear()
    '                Else
    '                    ComboBox3.SelectedText = dsFeatureManagement.Tables(0).Rows(0).Item("projectdimid")
    '                    'TextBox72.Text = isnull(dsProduct.Tables(0).Rows(0).Item("PM_Location"))
    '                End If


    '                'dsProductTickets = SQLHelper.ExecuteDataset(CN, "dbo.s_get_products_for_Products_tab", isnull(dsProduct.Tables(0).Rows(0).Item("Name")))
    '                'dgvProductsTickets_BindData()

    '            Else
    '                ' Clear_feature()
    '            End If

    '            bInitial = False
    '        End If

    '    Catch ex As Exception
    '        MsgBox("cmbfeatureMgmt_SelectedIndexChanged" + " : " + ex.Message)
    '        ' Functions.Sendmail(ex.Message, "cmbfeatureMgmt_SelectedIndexChanged" + " : " + Userid, 0, 0, "Project Management")
    '    End Try
    'End Sub
#End Region



End Class
