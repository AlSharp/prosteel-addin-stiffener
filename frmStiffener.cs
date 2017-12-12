/*--------------------------------------------------------------------------------------+
|
|     GSF - Garrison Steel Fabricators, Pell city, AL
|     Developer - Albert Sharapov, Steel Detailer
|
+--------------------------------------------------------------------------------------*/

using System;
using System.IO;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using System.Collections;
using System.Collections.Generic;
using SRI=System.Runtime.InteropServices;

using Bentley.ProStructures;
using ProStructuresOE;
using Bentley.ProStructures.CadSystem;
using Bentley.ProStructures.Configuration;
using Bentley.ProStructures.Drawing;
using Bentley.ProStructures.Geometry.Data;
using Bentley.ProStructures.Geometry.Utilities;
using Bentley.ProStructures.Miscellaneous;
using Bentley.ProStructures.Property;
using Bentley.ProStructures.Entity;
using Bentley.ProStructures.Steel.Plate;
using Bentley.ProStructures.Steel.Primitive;
using Bentley.ProStructures.Steel.Bolt;
using Bentley.ProStructures.Modification.Edit;
using Bentley.ProStructures.Modification.ObjectData;
using Microsoft.VisualBasic;
using PlugInBase;
using PlugInBase.PlugInCommon;

using BM=Bentley.MicroStation;
using BMW=Bentley.MicroStation.WinForms;
using BMI=Bentley.MicroStation.InteropServices;
using BCOM=Bentley.Interop.MicroStationDGN;

namespace GSFStiffener
{
    /// <summary>Main form for CellUtility AddIn GSFStiffener.
    /// </summary>
    public class GSFStiffenerControl : BMW.Adapter 
    {
        private static GSFStiffenerControl              s_current;
        private Bentley.Windowing.WindowContent  m_windowContent;

        private Bentley.MicroStation.AddIn                        m_addIn;

        private System.Windows.Forms.TabControl         tabControl1;

        //first tabpage
        private System.Windows.Forms.TabPage            stiffenerTabPage;

        private System.Windows.Forms.GroupBox 			stiffenerSettingsGroupBox;
        private System.Windows.Forms.Label 				stiffenerNameLabel;
        private System.Windows.Forms.Label              stiffenerMaterialLabel;
        private System.Windows.Forms.Label              stiffenerThicknessLabel;
        private System.Windows.Forms.Label              stiffenerPosNumLabel;
        private System.Windows.Forms.TextBox            stiffenerNameTextBox;
        private System.Windows.Forms.ComboBox           stiffenerMaterialComboBox;
        private System.Windows.Forms.ComboBox           stiffenerThicknessComboBox;
        private System.Windows.Forms.TextBox            stiffenerPosNumTextBox;

        private System.Windows.Forms.GroupBox           stiffenerPositionGroupBox;
        private System.Windows.Forms.Label              stiffenerPositionLabel;
        private System.Windows.Forms.ComboBox           stiffenerPositionComboBox;

        private System.Windows.Forms.Button             insertStiffenerOne;
        private System.Windows.Forms.Button             insertStiffenerMany;
        private System.Windows.Forms.Button             saveStiffenerToDatabase;

        private System.Windows.Forms.Button             okButton;
        private System.Windows.Forms.Button             cancelButton;
        private System.Windows.Forms.Button             helpButton;

        //second tabpage
        private System.Windows.Forms.TabPage            assignmentsTabPage;
        private System.Windows.Forms.Label              detailStyleLabel;
        private System.Windows.Forms.Label              displayClassLabel;
        private System.Windows.Forms.Label              areaClassLabel;
        private System.Windows.Forms.Label              partFamilyLabel;
        private System.Windows.Forms.Label              descriptionLabel;
        private System.Windows.Forms.Label              levelLabel;
        private System.Windows.Forms.ComboBox 			detailStyleComboBox;
        private System.Windows.Forms.ComboBox 			displayClassComboBox;
        private System.Windows.Forms.ComboBox           areaClassComboBox;
        private System.Windows.Forms.ComboBox           partFamilyComboBox;
        private System.Windows.Forms.ComboBox           descriptionComboBox;
        private System.Windows.Forms.ComboBox           levelComboBox;

        private System.Windows.Forms.Button             okButton1;
        private System.Windows.Forms.Button             cancelButton1;
        private System.Windows.Forms.Button             helpButton1;

        //tooltip
        private System.Windows.Forms.ToolTip 			toolTip1;

        //stores tabpageindex
        private static int                              tabPage;

        //stores objId for deletion if cancel operation
        private static int[]                            objIds;

        //stores values for stiffenerPositionComboBox
        private static string[]                         stiffenerPositions;

        /// <summary>The Visual Studio IDE requires a constructor without arguments.</summary>
        private GSFStiffenerControl()
        {
            System.Diagnostics.Debug.Assert (this.DesignMode, "Do not use the default constructor");
            InitializeComponent(tabPage);
        }
            
        /// <summary>Constructor</summary>
        internal GSFStiffenerControl(Bentley.MicroStation.AddIn addIn)
        {
            m_addIn     = addIn;
            InitializeComponent(tabPage);

            //  Set up events to handle resizing of form; closing of form
            this.Closed += new EventHandler(GSFStiffenerControl_Closed);
        }

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        protected override void Dispose( bool disposing )
        {
            base.Dispose( disposing );
        }

        #region Windows Form Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent(int TabIndex)
        {
            this.tabControl1 = new System.Windows.Forms.TabControl();

            this.stiffenerTabPage = new System.Windows.Forms.TabPage();

            this.stiffenerSettingsGroupBox = new System.Windows.Forms.GroupBox();

            this.stiffenerNameLabel = new System.Windows.Forms.Label();
            this.stiffenerMaterialLabel = new System.Windows.Forms.Label();
            this.stiffenerThicknessLabel = new System.Windows.Forms.Label();
            this.stiffenerPosNumLabel = new System.Windows.Forms.Label();

            this.stiffenerNameTextBox = new System.Windows.Forms.TextBox();
            this.stiffenerMaterialComboBox = new System.Windows.Forms.ComboBox();
            this.stiffenerThicknessComboBox = new System.Windows.Forms.ComboBox();
            this.stiffenerPosNumTextBox = new System.Windows.Forms.TextBox();

            this.stiffenerPositionGroupBox = new System.Windows.Forms.GroupBox();

            this.stiffenerPositionLabel = new System.Windows.Forms.Label();

            this.stiffenerPositionComboBox = new System.Windows.Forms.ComboBox();

            this.insertStiffenerOne = new System.Windows.Forms.Button();
            this.insertStiffenerMany = new System.Windows.Forms.Button();
            this.saveStiffenerToDatabase = new System.Windows.Forms.Button();

            this.okButton = new System.Windows.Forms.Button();
            this.cancelButton = new System.Windows.Forms.Button();
            this.helpButton = new System.Windows.Forms.Button();

            
            this.assignmentsTabPage = new System.Windows.Forms.TabPage();

            this.detailStyleLabel = new System.Windows.Forms.Label();
            this.displayClassLabel = new System.Windows.Forms.Label();
            this.areaClassLabel = new System.Windows.Forms.Label();
            this.partFamilyLabel = new System.Windows.Forms.Label();
            this.descriptionLabel = new System.Windows.Forms.Label();
            this.levelLabel = new System.Windows.Forms.Label();

            this.detailStyleComboBox = new System.Windows.Forms.ComboBox();
            this.displayClassComboBox = new System.Windows.Forms.ComboBox();
            this.areaClassComboBox = new System.Windows.Forms.ComboBox();
            this.partFamilyComboBox = new System.Windows.Forms.ComboBox();
            this.descriptionComboBox = new System.Windows.Forms.ComboBox();
            this.levelComboBox = new System.Windows.Forms.ComboBox();

            this.okButton1 = new System.Windows.Forms.Button();
            this.cancelButton1 = new System.Windows.Forms.Button();
            this.helpButton1 = new System.Windows.Forms.Button();
            
            
            this.toolTip1 = new System.Windows.Forms.ToolTip();

            this.SuspendLayout();

            //
            // Settings
            //
            //
            //  GroupBox settings
            //
            int space_bw_groupboxes = 8;
            //
            //Button settings
            //
            System.Drawing.Size buttonSize = new System.Drawing.Size (33, 33);
            int space_bw_buttons = 5;
            int space_bw_buttons_and_groupbox = 8;
            int step_buttonseries = buttonSize.Width + space_bw_buttons;
            //
            //Label settings
            //
            System.Drawing.Size labelSize = new System.Drawing.Size(100, 20);
            System.Drawing.Point firstPointlabelseries = new System.Drawing.Point(5, 20);
            int space_bw_labels = 5;
            int step_labelseries = labelSize.Height + space_bw_labels;
            System.Drawing.Size embedmentValuesLabelSize = new System.Drawing.Size(80, 40);
            //
            //Combobox settings
            //
            System.Drawing.Size comboboxSize = new System.Drawing.Size (130, 20);
            int space_bw_label_and_combobox = 1;
            System.Drawing.Point firstPointcomboboxseries = new System.Drawing.Point(
                firstPointlabelseries.X+ labelSize.Width + space_bw_label_and_combobox,
                firstPointlabelseries.Y);
            //
            //TextBox settings
            //
            System.Drawing.Size textboxSize = new System.Drawing.Size(130, 20);
            //
            //Locatition of first button
            //
            System.Drawing.Point firstPointbuttonseries = new System.Drawing.Point(
                12,
                firstPointlabelseries.Y + 8*labelSize.Height + 8*space_bw_labels +
                space_bw_buttons_and_groupbox + space_bw_groupboxes);
            //
            // Load combobox lists
            short num = -1;
            string configlevel = "";
            string defaultlevel = "";
            string actuallevel = "";
            //stiffener tabpage
            ListsFunctions.FillCombeWithTableValues(ref stiffenerThicknessComboBox, DescriptionTableType.kMainPlateThickTable);
            int material_list_int = ListsFunctions.DisplayMaterialList(ref this.stiffenerMaterialComboBox,ref num);
            stiffenerPositions = new string[]{
                "Right side", "Left side", "Both sides"
            };
            this.stiffenerPositionComboBox.Items.AddRange(stiffenerPositions);
            //assignments tabpage
            int det_styles_int = ListsFunctions.DisplayDetailStyleList(ref this.detailStyleComboBox, ref num);
            int display_classes_int = ListsFunctions.DisplayDisplayClassList(ref this.displayClassComboBox, ref num);
            int area_classed_int = ListsFunctions.DisplayAreaClassList(ref this.areaClassComboBox, ref num);
            int partfamily_classes_int = ListsFunctions.DisplayFamilyClassList(ref this.partFamilyComboBox, ref num);
            int description_classes_int = ListsFunctions.DisplayDescriptionClassList(ref this.descriptionComboBox, ref num);
            int level_int = ListsFunctions.DisplayLayerList(ref this.levelComboBox, ref defaultlevel, ref configlevel, ref actuallevel);


            //
            //  Load last saved values
            //
            LoadFromTemplate();   
        	// 
            //  Create the ToolTip
            //
            this.toolTip1.AutoPopDelay = 5000;
            this.toolTip1.InitialDelay = 1000;
            this.toolTip1.ReshowDelay = 500;
            // 
            //  Stiffener Settings
            //
            this.stiffenerSettingsGroupBox.Text = "Stiffener Settings";
            this.stiffenerSettingsGroupBox.Anchor = ((System.Windows.Forms.AnchorStyles)(
                (System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Top)));
            this.stiffenerSettingsGroupBox.Location = new System.Drawing.Point (8, 5);
            this.stiffenerSettingsGroupBox.Size = new System.Drawing.Size (
                2*firstPointlabelseries.X + labelSize.Width + space_bw_label_and_combobox + comboboxSize.Width,
                firstPointlabelseries.Y + 4*labelSize.Height + 4*space_bw_labels);
            //
            //  Position Settings
            //
            this.stiffenerPositionGroupBox.Text = "Position Settings";
            this.stiffenerPositionGroupBox.Anchor = ((System.Windows.Forms.AnchorStyles)(
                (System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Top)));
            this.stiffenerPositionGroupBox.Location = new System.Drawing.Point (
                8, 
                5 + this.stiffenerSettingsGroupBox.Size.Height + space_bw_groupboxes);
            this.stiffenerPositionGroupBox.Size = new System.Drawing.Size (
                this.stiffenerSettingsGroupBox.Size.Width,
                firstPointlabelseries.Y + 3*labelSize.Height + 3*space_bw_labels);
            //
            //  tabControl1
            //
            this.tabControl1.Location = new System.Drawing.Point(6, 7);
            this.tabControl1.Size = new System.Drawing.Size(
                2 + 3*this.stiffenerSettingsGroupBox.Location.X + this.stiffenerSettingsGroupBox.Size.Width,
                23 + firstPointbuttonseries.Y + buttonSize.Height + 5
                );
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.TabIndex = 0;
            //
            //stiffenerTabPage
            //
            this.stiffenerTabPage.Text = "Stiffener";
            this.stiffenerTabPage.TabIndex = 0;
            //
            //  assignmentsTabPage
            //
            this.assignmentsTabPage.Text = "Assignments";
            this.assignmentsTabPage.TabIndex = 1;
            //
            //  stiffenerTabPage Items:
            // 
            //
            //  Name label
            //
            this.stiffenerNameLabel.Text = "Label";
            this.stiffenerNameLabel.Size = labelSize;
            this.stiffenerNameLabel.Location = firstPointlabelseries;
            this.stiffenerNameLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            //
            //  Material label
            //
            this.stiffenerMaterialLabel.Text = "Material";
            this.stiffenerMaterialLabel.Size = labelSize;
            this.stiffenerMaterialLabel.Location = new System.Drawing.Point(
                firstPointlabelseries.X,
                firstPointlabelseries.Y + step_labelseries);
            this.stiffenerMaterialLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            //  Thickness label
            //
            this.stiffenerThicknessLabel.Text = "Thickness";
            this.stiffenerThicknessLabel.Size = labelSize;
            this.stiffenerThicknessLabel.Location = new System.Drawing.Point(
                firstPointlabelseries.X,
                firstPointlabelseries.Y + 2*step_labelseries);
            this.stiffenerThicknessLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            //  PosNum label
            //
            this.stiffenerPosNumLabel.Text = "Pos. number";
            this.stiffenerPosNumLabel.Size = labelSize;
            this.stiffenerPosNumLabel.Location = new System.Drawing.Point(
                firstPointlabelseries.X,
                firstPointlabelseries.Y + 3*step_labelseries);
            this.stiffenerPosNumLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            //  Name TextBox
            //
            this.stiffenerNameTextBox.Size = textboxSize;
            this.stiffenerNameTextBox.Location = firstPointcomboboxseries;
            //
            //  Material ComboBox
            //
            this.stiffenerMaterialComboBox.Size = comboboxSize;
            this.stiffenerMaterialComboBox.Location = new System.Drawing.Point(
                firstPointcomboboxseries.X,
                firstPointcomboboxseries.Y + step_labelseries);
            this.stiffenerThicknessComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            //
            //  Thickness ComboBox
            //
            this.stiffenerThicknessComboBox.Size = comboboxSize;
            this.stiffenerThicknessComboBox.Location = new System.Drawing.Point(
                firstPointcomboboxseries.X,
                firstPointcomboboxseries.Y + 2*step_labelseries);
            this.stiffenerThicknessComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            //
            //  PosNum TextBox
            //
            this.stiffenerPosNumTextBox.Size = textboxSize;
            this.stiffenerPosNumTextBox.Location = new System.Drawing.Point(
                firstPointcomboboxseries.X,
                firstPointcomboboxseries.Y + 3*step_labelseries);
            //
            //   Position Label
            //
            this.stiffenerPositionLabel.Text = "Position";
            this.stiffenerPositionLabel.Size = labelSize;
            this.stiffenerPositionLabel.Location = new System.Drawing.Point(
                firstPointlabelseries.X,
                firstPointlabelseries.Y);
            this.stiffenerPositionLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            //  Position ComboBox
            //
            this.stiffenerPositionComboBox.Size = comboboxSize;
            this.stiffenerPositionComboBox.Location = new System.Drawing.Point(
                firstPointcomboboxseries.X,
                firstPointcomboboxseries.Y);
            this.stiffenerPositionComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            //
            //  Assignment tabpage itmes
            //
            //
            //  Detail Style label
            //
            this.detailStyleLabel.Text = "Detail style";
            this.detailStyleLabel.Size = labelSize;
            this.detailStyleLabel.Location = firstPointlabelseries;
            this.detailStyleLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            //  Display Class label
            //
            this.displayClassLabel.Text = "Display class";
            this.displayClassLabel.Size = labelSize;
            this.displayClassLabel.Location = new System.Drawing.Point(
                firstPointlabelseries.X,
                firstPointlabelseries.Y + step_labelseries);
            this.displayClassLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            //  Area Class label
            //
            this.areaClassLabel.Text = "Area Class";
            this.areaClassLabel.Size = labelSize;
            this.areaClassLabel.Location = new System.Drawing.Point(
                firstPointlabelseries.X,
                firstPointlabelseries.Y + 2*step_labelseries);
            this.areaClassLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            //  Part Family label
            //
            this.partFamilyLabel.Text = "Part Family";
            this.partFamilyLabel.Size = labelSize;
            this.partFamilyLabel.Location = new System.Drawing.Point(
                firstPointlabelseries.X,
                firstPointlabelseries.Y + 3*step_labelseries);
            this.partFamilyLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            //  Description labe
            //
            this.descriptionLabel.Text = "Description";
            this.descriptionLabel.Size = labelSize;
            this.descriptionLabel.Location = new System.Drawing.Point(
                firstPointlabelseries.X,
                firstPointlabelseries.Y + 4*step_labelseries);
            this.descriptionLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            //  Lelel labe
            //
            this.levelLabel.Text = "Level";
            this.levelLabel.Size = labelSize;
            this.levelLabel.Location = new System.Drawing.Point(
                firstPointlabelseries.X,
                firstPointlabelseries.Y + 5*step_labelseries);
            this.levelLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            //
            //  Detail Styles ComboBox
            //
            this.detailStyleComboBox.Size = comboboxSize;
            this.detailStyleComboBox.Location = firstPointcomboboxseries;
            this.detailStyleComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            //
            //  Display Class ComboBox
            //
            this.displayClassComboBox.Size = comboboxSize;
            this.displayClassComboBox.Location = new System.Drawing.Point(
                firstPointcomboboxseries.X,
                firstPointcomboboxseries.Y + step_labelseries);
            this.displayClassComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            //
            //  Area Class ComboBox
            //
            this.areaClassComboBox.Size = comboboxSize;
            this.areaClassComboBox.Location = new System.Drawing.Point(
                firstPointcomboboxseries.X,
                firstPointcomboboxseries.Y + 2*step_labelseries);
            this.areaClassComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            //
            //  Part family ComboBox
            //
            this.partFamilyComboBox.Size = comboboxSize;
            this.partFamilyComboBox.Location = new System.Drawing.Point(
                firstPointcomboboxseries.X,
                firstPointcomboboxseries.Y + 3*step_labelseries);
            this.partFamilyComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            //
            //  Description ComboBox
            //
            this.descriptionComboBox.Size = comboboxSize;
            this.descriptionComboBox.Location = new System.Drawing.Point(
                firstPointcomboboxseries.X,
                firstPointcomboboxseries.Y + 4*step_labelseries);
            this.descriptionComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            //
            //  Level ComboBox
            //
            this.levelComboBox.Size = comboboxSize;
            this.levelComboBox.Location = new System.Drawing.Point(
                firstPointcomboboxseries.X,
                firstPointcomboboxseries.Y + 5*step_labelseries);
            this.levelComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;

            //
            // Load Bitmap Images
            //
            System.Reflection.Assembly ass1 = System.Reflection.Assembly.GetExecutingAssembly();
            Stream stream1 = ass1.GetManifestResourceStream("GSFStiffener.OkIcon.ico");
            Stream stream2 = ass1.GetManifestResourceStream("GSFStiffener.CancelIcon.ico");
            Stream stream3 = ass1.GetManifestResourceStream("GSFStiffener.InfoIcon.ico");
            // 
            //  insertStiffenerOne Button
            //
            this.insertStiffenerOne.Name = "insertStiffenerOne";
            // this.insertStiffenerOne.Image = Image.FromStream(stream1);
            this.toolTip1.SetToolTip(this.insertStiffenerOne, "Insert Stiffener");
            this.insertStiffenerOne.Size = buttonSize;
            this.insertStiffenerOne.Location = firstPointbuttonseries;
            this.insertStiffenerOne.Click += new System.EventHandler(this.cmdInsertStiffenerOne);
            //
            //  convertBolt
            //
            this.insertStiffenerMany.Name = "insertStiffenerMany";
            // this.insertStiffenerMany.Image = Image.FromStream(stream2);
            this.toolTip1.SetToolTip(this.insertStiffenerMany, "Create array of stiffeners");
            this.insertStiffenerMany.Size = buttonSize;
            this.insertStiffenerMany.Location = new System.Drawing.Point(
                firstPointbuttonseries.X + step_buttonseries,
                firstPointbuttonseries.Y);
            this.insertStiffenerMany.Click += new System.EventHandler(this.cmdInsertStiffenerMany);
            //
            //  OkButton
            //
            this.okButton.Name = "Ok";
            this.okButton.Image = Image.FromStream(stream1);
            this.toolTip1.SetToolTip(this.okButton, "OK");
            this.okButton.Size = buttonSize;
            this.okButton.Location = new System.Drawing.Point(
                firstPointbuttonseries.X + 3*step_buttonseries + 10,
                firstPointbuttonseries.Y);
            this.okButton.Click += new System.EventHandler(this.cmdOk);
            //
            //  CancelButton
            //
            this.cancelButton.Name = "Cancel";
            this.cancelButton.Image = Image.FromStream(stream2);
            this.toolTip1.SetToolTip(this.cancelButton, "Cancel");
            this.cancelButton.Size = buttonSize;
            this.cancelButton.Location = new System.Drawing.Point(
                firstPointbuttonseries.X + 4*step_buttonseries + 10,
                firstPointbuttonseries.Y);
            this.cancelButton.Click += new System.EventHandler(this.cmdCancel);
            //
            //  helpBolt
            //
            this.helpButton.Name = "Help";
            this.helpButton.Image = Image.FromStream(stream3);
            this.toolTip1.SetToolTip(this.helpButton, "About");
            this.helpButton.Size = buttonSize;
            this.helpButton.Location = new System.Drawing.Point(
                firstPointbuttonseries.X + 5*step_buttonseries + 10,
                firstPointbuttonseries.Y);
            this.helpButton.Click += new System.EventHandler(this.cmdHelp);
            //
            //  OkButton
            //
            this.okButton1.Name = "Ok";
            this.okButton1.Image = Image.FromStream(stream1);
            this.toolTip1.SetToolTip(this.okButton1, "OK");
            this.okButton1.Size = buttonSize;
            this.okButton1.Location = new System.Drawing.Point(
                firstPointbuttonseries.X + 3*step_buttonseries + 10,
                firstPointbuttonseries.Y);
            this.okButton1.Click += new System.EventHandler(this.cmdOk);
            //
            //  CancelButton1
            //
            this.cancelButton1.Name = "Cancel";
            this.cancelButton1.Image = Image.FromStream(stream2);
            this.toolTip1.SetToolTip(this.cancelButton1, "Cancel");
            this.cancelButton1.Size = buttonSize;
            this.cancelButton1.Location = new System.Drawing.Point(
                firstPointbuttonseries.X + 4*step_buttonseries + 10,
                firstPointbuttonseries.Y);
            this.cancelButton1.Click += new System.EventHandler(this.cmdCancel);
            //
            //  helpBolt
            //
            this.helpButton1.Name = "Help";
            this.helpButton1.Image = Image.FromStream(stream3);
            this.toolTip1.SetToolTip(this.helpButton1, "About");
            this.helpButton1.Size = buttonSize;
            this.helpButton1.Location = new System.Drawing.Point(
                firstPointbuttonseries.X + 5*step_buttonseries + 10,
                firstPointbuttonseries.Y);
            this.helpButton1.Click += new System.EventHandler(this.cmdHelp);
            //
            //  Stiffener Setting GroupBox Controls
            //
            this.stiffenerSettingsGroupBox.Controls.Add(this.stiffenerNameLabel);
            this.stiffenerSettingsGroupBox.Controls.Add(this.stiffenerMaterialLabel);
            this.stiffenerSettingsGroupBox.Controls.Add(this.stiffenerThicknessLabel);
            this.stiffenerSettingsGroupBox.Controls.Add(this.stiffenerPosNumLabel);
            this.stiffenerSettingsGroupBox.Controls.Add(this.stiffenerNameTextBox);
            this.stiffenerSettingsGroupBox.Controls.Add(this.stiffenerMaterialComboBox);
            this.stiffenerSettingsGroupBox.Controls.Add(this.stiffenerThicknessComboBox);
            this.stiffenerSettingsGroupBox.Controls.Add(this.stiffenerPosNumTextBox);
            //
            //  Stiffener Position Settings Groupbox Controls
            //
            this.stiffenerPositionGroupBox.Controls.Add(this.stiffenerPositionLabel);
            this.stiffenerPositionGroupBox.Controls.Add(this.stiffenerPositionComboBox);
            //
            //  Stiffener TabPage Controls
            //
            this.stiffenerTabPage.Controls.Add(this.stiffenerSettingsGroupBox);
            this.stiffenerTabPage.Controls.Add(this.stiffenerPositionGroupBox);
            this.stiffenerTabPage.Controls.Add(this.insertStiffenerOne);
            this.stiffenerTabPage.Controls.Add(this.insertStiffenerMany);
            this.stiffenerTabPage.Controls.Add(this.okButton);
            this.stiffenerTabPage.Controls.Add(this.cancelButton);
            this.stiffenerTabPage.Controls.Add(this.helpButton);
            //
            // TabPage expAnchorSettings Controls
            //
            this.assignmentsTabPage.Controls.Add(this.detailStyleLabel);
            this.assignmentsTabPage.Controls.Add(this.displayClassLabel);
            this.assignmentsTabPage.Controls.Add(this.areaClassLabel);
            this.assignmentsTabPage.Controls.Add(this.partFamilyLabel);
            this.assignmentsTabPage.Controls.Add(this.descriptionLabel);
            this.assignmentsTabPage.Controls.Add(this.levelLabel);

            this.assignmentsTabPage.Controls.Add(this.detailStyleComboBox);
            this.assignmentsTabPage.Controls.Add(this.displayClassComboBox);
            this.assignmentsTabPage.Controls.Add(this.areaClassComboBox);
            this.assignmentsTabPage.Controls.Add(this.partFamilyComboBox);
            this.assignmentsTabPage.Controls.Add(this.descriptionComboBox);
            this.assignmentsTabPage.Controls.Add(this.levelComboBox);


            this.assignmentsTabPage.Controls.Add(this.okButton1);
            this.assignmentsTabPage.Controls.Add(this.cancelButton1);
            this.assignmentsTabPage.Controls.Add(this.helpButton1);
            //
            //TabControl
            //
            this.tabControl1.Controls.Add(this.stiffenerTabPage);
            this.tabControl1.Controls.Add(this.assignmentsTabPage);
            //opens tabpage with index = TabPage
            this.tabControl1.SelectedIndex = TabIndex;
            //
            //Add control to the form
            //
            this.Controls.Add(this.tabControl1);
            this.Name = "ProSteel AddIn";
            this.Text = "Stiffener";
            this.ResumeLayout(false);
        }
            
        #endregion
        /// <summary>
        /// Show the form if it is not already displayed
        /// </summary>
        internal static void ShowForm (Bentley.MicroStation.AddIn addIn)
        {
            if (null != s_current)
                return;

            s_current = new GSFStiffenerControl (addIn);
            s_current.AttachAsTopLevelForm (addIn, true);

            
            s_current.AutoOpen = true;
            s_current.AutoOpenKeyin = "mdl load GSFStiffener";
            
            s_current.NETDockable = false; 
            Bentley.Windowing.WindowManager    windowManager = 
                        Bentley.Windowing.WindowManager.GetForMicroStation ();
            s_current.m_windowContent = 
                windowManager.DockPanel (s_current, s_current.Name, s_current.Text, 
                Bentley.Windowing.DockLocation.Floating);
            AdapterWorkarounds.WCFixedBorder.SetBorderFixed(s_current.m_windowContent);

            s_current.m_windowContent.CanDockHorizontally = false; // limit to left and right docking
            s_current.m_windowContent.CanDockVertically = false;
        }
        /// <summary>
        /// Close the form if it is currently displayed
        /// </summary>
        internal static void CloseForm ()
        {
            if (s_current == null)
                return;
            s_current.m_windowContent.Close();

            s_current = null;
        }
        /// <summary>Handle the standard Closed event
        /// </summary>
        private void GSFStiffenerControl_Closed(object sender, EventArgs e)
        {   
            if (s_current != null)
            tabPage = this.tabControl1.SelectedIndex; //stores last opened tabpage
            s_current.Dispose (true);
              
            s_current = null;
        }
        //
        //Save field values into template
        //
        private void SaveInToTemplate()
        {
            PsTemplateManager psTemplateManager = new PsTemplateManager();
            psTemplateManager.AppendNumber(1);
            psTemplateManager.AppendString(this.stiffenerNameTextBox.Text);
            psTemplateManager.AppendString(this.stiffenerMaterialComboBox.Text.ToString());
            psTemplateManager.AppendString(this.stiffenerThicknessComboBox.Text.ToString());
            psTemplateManager.AppendString(this.stiffenerPositionComboBox.Text.ToString());
            psTemplateManager.SaveTemplateEntry("GSFStiffener", 1);
        }
        //
        //Load field values from template
        //
        private void LoadFromTemplate()
        {
            PsTemplateManager psTemplateManager = new PsTemplateManager();
            psTemplateManager.LoadTemplateEntry("GSFStiffener");
            if (psTemplateManager.IsLoaded)
            {
                long num1 = (long)psTemplateManager.get_Number(0);
                this.stiffenerNameTextBox.Text = psTemplateManager.get_String(0);
                this.stiffenerMaterialComboBox.Text = psTemplateManager.get_String(1);
                this.stiffenerThicknessComboBox.Text = psTemplateManager.get_String(2);
                this.stiffenerPositionComboBox.Text = psTemplateManager.get_String(3);
            }
            else
            {
                this.stiffenerNameTextBox.Text = "PL";
                this.stiffenerMaterialComboBox.Text = "ASTM A36 Gr.36";
                this.stiffenerThicknessComboBox.Text = "0:0 3/8";
            }
        }
        private void cmdInsertStiffenerOne(object sender, System.EventArgs e)
        {
            PsUnits psUnits = new PsUnits();
            string stiffenerLabel = this.stiffenerNameTextBox.Text;
            string stiffenerMaterial = this.stiffenerMaterialComboBox.Text;
            double stiffenerThickness = psUnits.ConvertToNumeric(this.stiffenerThicknessComboBox.Text);
            int positionIndex = this.stiffenerPositionComboBox.SelectedIndex;
            SaveInToTemplate();
            GSFStiffenerControl.CloseForm();
            InsertOne(stiffenerLabel, stiffenerMaterial, stiffenerThickness, positionIndex);
            GSFStiffenerControl.ShowForm(m_addIn);
        }
        private void InsertOne(string StiffenerLabel, string StiffenerMaterial, double StiffenerThickness, int StiffenerPositionIndex)
        {
            List<int> objs = new List<int>();
            PsUnits psUnits = new PsUnits();
            PsShapeInfo psShapeInfo = new PsShapeInfo();
            PsGeometryFunctions geometryFunctions = new PsGeometryFunctions();
            PsPoint psInsertionPoint = new PsPoint();
            PsPoint psStartPoint = new PsPoint();
            PsPoint psEndPoint = new PsPoint();
            PsPoint psMidLineInsertionPoint = new PsPoint();
            PsPolygon psStiffenerPolygon = new PsPolygon();
            PsCreatePlate psStiffinerPlateRight = new PsCreatePlate();
            PsCreatePlate psStiffinerPlateLeft = new PsCreatePlate();
            PsMatrix psMatrixRight = new PsMatrix();
            PsMatrix psMatrixLeft = new PsMatrix();
        	PsSelection arg1 = new PsSelection();
            arg1.SetSelectionFilter(SelectionFilter.kFilterShape);
            int num1 = arg1.PickObject("Please select W shape beam..."); //returns object id
            if (num1 > 0)
            {
                PsSelection arg2 = new PsSelection();
                int num2 = arg2.PickPoint("Select insertion point", CoordSystem.kWcs, psInsertionPoint);
                if (num2 ==1)
                {
                    psShapeInfo.SetObjectId(num1);
                    int num3 = psShapeInfo.GetInfo(); //gets object property values
                    psStartPoint = psShapeInfo.MidLineStart;
                    psEndPoint = psShapeInfo.MidLineEnd;
                    psMidLineInsertionPoint = geometryFunctions.OrthoProjectPointToLine(
                        psInsertionPoint,
                        psShapeInfo.MidLineStart,
                        psShapeInfo.MidLineEnd);
                    double shapeLength = geometryFunctions.GetDistanceBetween(psStartPoint, psEndPoint);
                    if (geometryFunctions.GetDistanceBetween(psMidLineInsertionPoint, psStartPoint) <= shapeLength &&
                        geometryFunctions.GetDistanceBetween(psMidLineInsertionPoint, psEndPoint) <= shapeLength)
                    {
                        psShapeInfo.SetStiffenerType(0, 0.75, 0);
                        psStiffenerPolygon = psShapeInfo.getStiffenerPolygon(psInsertionPoint);
                        psStiffinerPlateRight.SetFromPolygon(psStiffenerPolygon);
                        psStiffinerPlateRight.SetThickness(StiffenerThickness);
                        psMatrixRight = psShapeInfo.MidLineUcs;
                        psMatrixLeft = psMatrixRight.Clone();
                        psMatrixLeft.MirrorAtLine(psStartPoint, psEndPoint);

                        psStiffinerPlateRight.SetInsertMatrix(psMatrixRight);
                        psStiffinerPlateRight.Create();
                        objs.Add(psStiffinerPlateRight.ObjectId);

                        CommonFunctions.MoveObject(
                            (long)psStiffinerPlateRight.ObjectId,
                            ref psStartPoint,
                            ref psMidLineInsertionPoint);
                        
                        psStiffinerPlateLeft = psStiffinerPlateRight;
                        psStiffinerPlateLeft.SetInsertMatrix(psMatrixLeft);
                        psStiffinerPlateLeft.Create();
                        objs.Add(psStiffinerPlateLeft.ObjectId);

                        CommonFunctions.MoveObject(
                            (long)psStiffinerPlateLeft.ObjectId,
                            ref psStartPoint,
                            ref psMidLineInsertionPoint);

                        objIds = objs.ToArray();

                        //test
                        MessageBox.Show(psShapeInfo.ShapeName); 
                    }
                    else
                    {
                        MessageBox.Show("Selected point is out of beam", "Oops...", MessageBoxButtons.OK);
                    }        
                }
            }
        }
        private void cmdInsertStiffenerMany(object sender, System.EventArgs e)
    	{
            string str1 = this.stiffenerPosNumTextBox.Text;
    		GSFStiffenerControl.CloseForm();
    		InsertMany(str1);
    		GSFStiffenerControl.ShowForm(m_addIn);
    	}
        private void InsertMany(string type)
    	{
            MessageBox.Show(type);
            var result = IBeamDb.GetIBeamByType(type);
    		// this.stiffenerNameTextBox.Text = st1;
            if (result != null)
            {
                MessageBox.Show("YES " + result._id.ToString() + " " + result.Assigned.ToString());  
            }
            else
            {
                MessageBox.Show("NO " + result._id.ToString()); 
            }
    	}
        private void cmdHelp(object sender, System.EventArgs e)
        {
            Help();
        }
        private void Help()
    	{
    		string message = 
    		"ProSteel Addin GSF TC Bolt/Exp. Anchor v1.1" + Environment.NewLine +
    		"1. Creates tension control bolts" + Environment.NewLine +
    		"2. Replaces standard bolts with tension control bolts" + Environment.NewLine +
    		"3. Create expansion anchors with custom name" + Environment.NewLine +
    		"-------------------------------------------------------" + Environment.NewLine +
    		"Developed by Steel Detailer Albert E Sharapov" + Environment.NewLine +
    		"Copyright \u00A9 2017 Garrison Steel Fabricators, Inc." + Environment.NewLine +
    		"All Rights Reserved.";
    		string caption = "About GSF TC Bolt/Exp. Anchor";
    		DialogResult result;
    		result = MessageBox.Show(message, caption, MessageBoxButtons.OK, MessageBoxIcon.Information);
    	}
        private void cmdOk(object sender, System.EventArgs e)
        {
            Ok();
        }
        private void Ok()
    	{
    		GSFStiffenerControl.CloseForm();
    	}
        private void cmdCancel(object sender, System.EventArgs e)
        {
            Cancel();
        }
        private void Cancel()
    	{
    		// PsTransaction psTransaction = new PsTransaction();
    		// psTransaction.EraseLongId((long)objId);
    		// psTransaction.Close();
            if (!objIds)
            {
                foreach (int objId in objIds)
                {
                    long i = (long)objId;
                    CommonFunctions.EntityDelete(ref i);
                }  
            }    
    		GSFStiffenerControl.CloseForm();
    	}
    }   
}   
