﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using ExceptionReporting.Core;
using ExceptionReporting.Extensions;
using ExceptionReporting.SystemInfo;

#pragma warning disable 1591

namespace ExceptionReporting.Views
{
	/// <summary>
	/// The main ExceptionReporter dialog
	/// </summary>
	// ReSharper disable UnusedMember.Global
	public partial class ExceptionReportView : Form, IExceptionReportView
	// ReSharper restore UnusedMember.Global
	{
		private bool _isDataRefreshRequired;
		private readonly ExceptionReportPresenter _presenter;

		public ExceptionReportView(ExceptionReportInfo reportInfo)
		{
			ShowFullDetail = true;
			InitializeComponent();
			TopMost = reportInfo.TopMost;

			_presenter = new ExceptionReportPresenter(this, reportInfo);

			WireUpEvents();
			PopulateTabs();
			PopulateReportInfo(reportInfo);
		}

		private void PopulateReportInfo(ExceptionReportInfo reportInfo)
		{
			urlEmail.Text = reportInfo.ContactEmail;
			txtFax.Text = reportInfo.Fax;
			lblContactMessageTop.Text = reportInfo.ContactMessageTop;
			txtPhone.Text = reportInfo.Phone;
			urlWeb.Text = reportInfo.WebUrl;
			lblExplanation.Text = reportInfo.UserExplanationLabel;
			ShowFullDetail = reportInfo.ShowFullDetail;
			ToggleShowFullDetail();
			btnDetailToggle.Visible = reportInfo.ShowLessMoreDetailButton;

			//TODO: show all exception messages
			txtExceptionMessageLarge.Text =
					txtExceptionMessage.Text =
					!string.IsNullOrEmpty(reportInfo.CustomMessage) ? reportInfo.CustomMessage : reportInfo.Exceptions[0].Message;

			txtExceptionMessageLarge2.Text = txtExceptionMessageLarge.Text;

			txtDate.Text = reportInfo.ExceptionDate.ToShortDateString();
			txtTime.Text = reportInfo.ExceptionDate.ToShortTimeString();
			txtUserName.Text = reportInfo.UserName;
			txtMachine.Text = reportInfo.MachineName;
			txtRegion.Text = reportInfo.RegionInfo;
			txtApplicationName.Text = reportInfo.AppName;
			txtVersion.Text = reportInfo.AppVersion;

			btnClose.FlatStyle =
					btnDetailToggle.FlatStyle =
					btnCopy.FlatStyle =
					btnEmail.FlatStyle =
					btnSave.FlatStyle = (reportInfo.ShowFlatButtons ? FlatStyle.Flat : FlatStyle.Standard);

			listviewAssemblies.BackColor =
					txtFax.BackColor =
					txtMachine.BackColor =
					txtPhone.BackColor =
					txtRegion.BackColor =
					txtTime.BackColor =
					txtTime.BackColor =
					txtUserName.BackColor =
					txtVersion.BackColor =
					txtApplicationName.BackColor =
					txtDate.BackColor =
					txtExceptionMessageLarge.BackColor =
					txtExceptionMessage.BackColor = reportInfo.BackgroundColor;

			if (!reportInfo.ShowButtonIcons)
			{
				RemoveButtonIcons();
			}

			Text = reportInfo.TitleText;
			txtUserExplanation.Font = new Font(txtUserExplanation.Font.FontFamily, reportInfo.UserExplanationFontSize);
			lblContactCompany.Text = string.Format("If this problem persists, please contact {0} support.", reportInfo.CompanyName);
			if (!reportInfo.CompanyName.IsEmpty())
			{
				btnSimpleEmail.Text = string.Format("Email {0}", reportInfo.CompanyName);
			}

			if (reportInfo.TakeScreenshot)
			{
				try
				{
					reportInfo.ScreenshotImage = ScreenshotTaker.TakeScreenShot();
				}
				catch { }
				// not too concerned about the specifics at the moment, just that an exception here doesn't prevent the entire mechansim from working
				// specifically, if we are raising this exception as the result of an out-of-memory exception, we have little chance of a screenshot succeeding
			}
		}

		private void RemoveButtonIcons()
		{
			// removing the icons, requires a bit of reshuffling of positions
			btnCopy.Image = btnEmail.Image = btnSave.Image = null;
			btnClose.Height = btnDetailToggle.Height = btnCopy.Height = btnEmail.Height = btnSave.Height = 27;
			btnClose.TextAlign = btnDetailToggle.TextAlign = btnCopy.TextAlign = btnEmail.TextAlign = btnSave.TextAlign = ContentAlignment.MiddleCenter;
			btnClose.Font = btnDetailToggle.Font = btnSave.Font = btnEmail.Font = btnCopy.Font = new Font(btnCopy.Font.FontFamily, 8.25f);
			ShiftDown3Pixels(new[] { btnClose, btnDetailToggle, btnCopy, btnEmail, btnSave });
		}

		private static void ShiftDown3Pixels(IEnumerable<Control> buttons)
		{
			foreach (var button in buttons)
			{
				button.Location = Point.Add(button.Location, new Size(new Point(0, 3)));
			}
		}

		private void WireUpEvents()
		{
			btnEmail.Click += Email_Click;
			btnSimpleEmail.Click += Email_Click;
			btnCopy.Click += Copy_Click;
			btnSimpleCopy.Click += Copy_Click;
			btnClose.Click += Close_Click;
			btnDetailToggle.Click += Detail_Click;
			btnSimpleDetailToggle.Click += Detail_Click;
			urlEmail.LinkClicked += EmailLink_Clicked;
			btnSave.Click += Save_Click;
			urlWeb.LinkClicked += UrlLink_Clicked;
			KeyPreview = true;
			KeyDown += ExceptionReportView_KeyDown;
		}

		private void ExceptionReportView_KeyDown(object sender, KeyEventArgs e)
		{
			if (e.KeyCode == Keys.Escape)
			{
				Close();
			}
		}

		public string ProgressMessage
		{
			set
			{
				lblProgressMessage.Visible = true;
				lblProgressMessage.Text = value;
			}
		}

		public bool EnableEmailButton
		{
			set { btnEmail.Enabled = value; }
		}

		public bool ShowProgressBar
		{
			set { progressBar.Visible = value; }
		}

		private bool ShowProgressLabel
		{
			set { lblProgressMessage.Visible = value; }
		}

		public bool ShowFullDetail { get; set; }

		public void ToggleShowFullDetail()
		{
			if (ShowFullDetail)
			{
				Size = new Size(625, 456);
				lessDetailPanel.Hide();
				btnDetailToggle.Text = "Less Detail";
				tabControl.Visible = true;
			}
			else
			{
				Size = new Size(400, 240);
				lessDetailPanel.Show();
				btnDetailToggle.Text = "More Detail";
				tabControl.Visible = false;
			}
		}

		public string UserExplanation
		{
			get { return txtUserExplanation.Text; }
		}

		public void Completed(bool success)
		{
			ProgressMessage = success ? "Email sent" : "Failed to send Email";
			ShowProgressBar = false;
			btnEmail.Enabled = true;
		}

		public void SetEmailCompletedState_WithMessageIfSuccess(bool success, string successMessage)
		{
			Completed(success);

			if (success)
			{
				ProgressMessage = successMessage;
			}
		}

		/// <summary>
		/// starts with all tabs visible, and removes ones that aren't configured to show
		/// </summary>
		private void PopulateTabs()
		{
			if (!_presenter.ReportInfo.ShowGeneralTab)
			{
				tabControl.TabPages.Remove(tabGeneral);
			}
			if (!_presenter.ReportInfo.ShowExceptionsTab)
			{
				tabControl.TabPages.Remove(tabExceptions);
			}
			if (!_presenter.ReportInfo.ShowAssembliesTab)
			{
				tabControl.TabPages.Remove(tabAssemblies);
			}
			if (!_presenter.ReportInfo.ShowSysInfoTab)
			{
				tabControl.TabPages.Remove(tabSysInfo);
			}
			if (!_presenter.ReportInfo.ShowContactTab)
			{
				tabControl.TabPages.Remove(tabContact);
			}
		}

		//TODO consider putting on a background thread - and avoid the OnActivated event altogether
		protected override void OnActivated(EventArgs e)
		{
			base.OnActivated(e);

			if (_isDataRefreshRequired)
			{
				_isDataRefreshRequired = false;
				_presenter.PopulateReport();
			}
		}

		public void SetProgressCompleteState()
		{
			Cursor = Cursors.Default;
			ShowProgressLabel = ShowProgressBar = false;
		}

		public void ShowExceptionReport()
		{
			_isDataRefreshRequired = true;
			ShowDialog();
		}

		public void SetInProgressState()
		{
			Cursor = Cursors.WaitCursor;
			ShowProgressLabel = true;
			ShowProgressBar = true;
			Application.DoEvents();
		}

		public void PopulateExceptionTab(IList<Exception> exceptions)
		{
			if (exceptions.Count == 1)
			{
				var exception = exceptions[0];
				AddExceptionControl(tabExceptions, exception);
			}
			else
			{
				var innerTabControl = new TabControl { Dock = DockStyle.Fill };
				tabExceptions.Controls.Add(innerTabControl);
				for (var index = 0; index < exceptions.Count; index++)
				{
					var exception = exceptions[index];
					var tabPage = new TabPage { Text = string.Format("Excepton {0}", index + 1) };
					innerTabControl.TabPages.Add(tabPage);
					AddExceptionControl(tabPage, exception);
				}
			}
		}

		private void AddExceptionControl(Control control, Exception exception)
		{
			var exceptionDetail = new ExceptionDetailControl();
			exceptionDetail.SetControlBackgrounds(_presenter.ReportInfo.BackgroundColor);
			exceptionDetail.PopulateExceptionTab(exception);
			exceptionDetail.Dock = DockStyle.Fill;
			control.Controls.Add(exceptionDetail);
		}

		public void PopulateAssembliesTab()
		{
			listviewAssemblies.Clear();
			listviewAssemblies.Columns.Add("Name", 320, HorizontalAlignment.Left);
			listviewAssemblies.Columns.Add("Version", 150, HorizontalAlignment.Left);

			var assemblies = new List<AssemblyName>(_presenter.AppAssembly.GetReferencedAssemblies())
																 {
																		 _presenter.AppAssembly.GetName()
																 };
			assemblies.Sort((x, y) => string.CompareOrdinal(x.Name, y.Name));
			foreach (var assemblyName in assemblies)
			{
				AddAssembly(assemblyName);
			}
		}

		private void AddAssembly(AssemblyName assemblyName)
		{
			var listViewItem = new ListViewItem { Text = assemblyName.Name };
			listViewItem.SubItems.Add(assemblyName.Version.ToString());
			listviewAssemblies.Items.Add(listViewItem);
		}

		protected override void OnClosing(CancelEventArgs e)
		{
			_presenter.Close();
			base.OnClosing(e);
		}

		private TreeNode CreateSysInfoTree()
		{
			var rootNode = new TreeNode("System");

			foreach (var sysInfoResult in _presenter.GetSysInfoResults())
			{
				SysInfoResultMapper.AddTreeViewNode(rootNode, sysInfoResult);
			}

			return rootNode;
		}

		public void PopulateSysInfoTab()
		{
			var rootNode = CreateSysInfoTree();
			treeEnvironment.Nodes.Add(rootNode);
			rootNode.Expand();
		}

		private void Copy_Click(object sender, EventArgs e)
		{
			_presenter.CopyReportToClipboard();
		}

		private void Close_Click(object sender, EventArgs e)
		{
			Close();
		}

		private void Detail_Click(object sender, EventArgs e)
		{
			_presenter.ToggleDetail();
		}

		private void Email_Click(object sender, EventArgs e)
		{
			_presenter.SendReportByEmail();
		}

		private void Save_Click(object sender, EventArgs e)
		{
			var saveDialog = new SaveFileDialog
			{
				Filter = "Text Files (*.txt)|*.txt|All files (*.*)|*.*",
				FilterIndex = 1,
				RestoreDirectory = true
			};

			if (saveDialog.ShowDialog() == DialogResult.OK)
			{
				_presenter.SaveReportToFile(saveDialog.FileName);
			}
		}


		private void UrlLink_Clicked(object sender, LinkLabelLinkClickedEventArgs e)
		{
			_presenter.NavigateToWebsite();
		}

		private void EmailLink_Clicked(object sender, LinkLabelLinkClickedEventArgs e)
		{
			_presenter.SendContactEmail();
		}

		public void ShowError(string message, Exception exception)
		{
			MessageBox.Show(message);       // last resort, hope it never happens
		}

	}
}

#pragma warning restore 1591