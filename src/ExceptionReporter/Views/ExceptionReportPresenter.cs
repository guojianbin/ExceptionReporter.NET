using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Text;
using ExceptionReporting.Core;
using ExceptionReporting.Mail;
using ExceptionReporting.SystemInfo;

namespace ExceptionReporting.Views
{
	/// <summary>
	/// The Presenter in this MVP (Model-View-Presenter) implementation 
	/// </summary>
	public class ExceptionReportPresenter
	{
		private readonly IExceptionReportView _view;
		private readonly ExceptionReportGenerator _reportGenerator;

		/// <summary>
		/// 
		/// </summary>
		/// <param name="view"></param>
		/// <param name="info"></param>
		public ExceptionReportPresenter(IExceptionReportView view, ExceptionReportInfo info)
		{
			_view = view;
			ReportInfo = info;
			_reportGenerator = new ExceptionReportGenerator(info);
		}

		/// <summary>
		/// The application assembly - ie the main application using the exception reporter assembly
		/// </summary>
		public Assembly AppAssembly
		{
			get { return ReportInfo.AppAssembly; }
		}

		/// <summary>
		/// 
		/// </summary>
		public ExceptionReportInfo ReportInfo { get; private set; }
		private WinFormsClipboard _clipboard = new WinFormsClipboard();

		private ExceptionReport CreateExceptionReport()
		{
			ReportInfo.UserExplanation = _view.UserExplanation;
			return _reportGenerator.CreateExceptionReport();
		}

		/// <summary>
		/// Save the exception report to file/disk
		/// </summary>
		/// <param name="fileName">the filename to save</param>
		public void SaveReportToFile(string fileName)
		{
			if (string.IsNullOrEmpty(fileName)) return;

			var exceptionReport = CreateExceptionReport();

			try
			{
				using (var stream = File.OpenWrite(fileName))
				{
					var writer = new StreamWriter(stream);
					writer.Write(exceptionReport);
					writer.Flush();
				}
			}
			catch (Exception exception)
			{
				_view.ShowError(string.Format("Unable to save file '{0}'", fileName), exception);
			}
		}

		/// <summary>
		/// Send the exception report via email, using the configured email method/type
		/// </summary>
		public void SendReportByEmail()
		{
			if (ReportInfo.MailMethod == ExceptionReportInfo.EmailMethod.SimpleMAPI)
			{
				SendMapiEmail();
			}

			if (ReportInfo.MailMethod == ExceptionReportInfo.EmailMethod.SMTP)
			{
				SendSmtpMailAsync();
			}
		}

		/// <summary>
		/// copy the report to the clipboard, using the clipboard implementation supplied
		/// </summary>
		public void CopyReportToClipboard()
		{
			var exceptionReport = CreateExceptionReport();
			_clipboard.CopyTo(exceptionReport.ToString());
			_view.ProgressMessage = string.Format("{0} copied to clipboard", ReportInfo.TitleText);
		}

		/// <summary>
		/// toggle the detail between 'simple' (just message) and showFullDetail (ie normal)
		/// </summary>
		public void ToggleDetail()
		{
			_view.ShowFullDetail = !_view.ShowFullDetail;
			_view.ToggleShowFullDetail();
		}

		private string BuildEmailText()
		{
			var emailTextBuilder = new EmailTextBuilder();
			var emailIntroString = emailTextBuilder.CreateIntro(ReportInfo.TakeScreenshot);
			var entireEmailText = new StringBuilder(emailIntroString);

			var report = CreateExceptionReport();
			entireEmailText.Append(report);

			return entireEmailText.ToString();
		}

		private async void SendSmtpMailAsync()
		{
			var emailText = BuildEmailText();

			_view.ProgressMessage = "Sending email via SMTP...";
			_view.EnableEmailButton = false;
			_view.ShowProgressBar = true;

			try
			{
				var mailSender = new MailSender(ReportInfo);
				await mailSender.SendSmtpAsync(emailText, _view);
			}
			catch (Exception exception)
			{
				_view.Completed(false);
				_view.ShowError("Unable to send email using SMTP" + Environment.NewLine + exception.Message, exception);
			}
		}

		private void SendMapiEmail()
		{
			var emailText = BuildEmailText();

			_view.ProgressMessage = "Launching email program...";
			_view.EnableEmailButton = false;

			var wasSuccessful = false;

			try
			{
				var mailSender = new MailSender(ReportInfo);
				mailSender.SendMapi(emailText);
				wasSuccessful = true;
			}
			catch (Exception exception)
			{
				wasSuccessful = false;
				_view.ShowError("Unable to send Email using 'Simple MAPI'", exception);
			}
			finally
			{
				_view.SetEmailCompletedState_WithMessageIfSuccess(wasSuccessful, string.Empty);
			}
		}

		/// <summary>
		/// Fetch the WMI system information
		/// </summary>
		public IEnumerable<SysInfoResult> GetSysInfoResults()
		{
			return _reportGenerator.GetOrFetchSysInfoResults();
		}

		/// <summary>
		/// Send email (using ShellExecute) to the configured contact email address
		/// </summary>
		public void SendContactEmail()
		{
			ShellExecute(string.Format("mailto:{0}", ReportInfo.ContactEmail));
		}

		/// <summary>
		/// Navigate to the website configured
		/// </summary>
		public void NavigateToWebsite()
		{
			ShellExecute(ReportInfo.WebUrl);
		}

		private void ShellExecute(string executeString)
		{
			try
			{
				var psi = new ProcessStartInfo(executeString) { UseShellExecute = true };
				Process.Start(psi);
			}
			catch (Exception exception)
			{
				_view.ShowError(string.Format("Unable to (Shell) Execute '{0}'", executeString), exception);
			}
		}

		/// <summary>
		/// The main entry point, populates the report with everything it needs
		/// </summary>
		public void PopulateReport()
		{
			try
			{
				_view.SetInProgressState();

				_view.PopulateExceptionTab(ReportInfo.Exceptions);
				_view.PopulateAssembliesTab();
				if (!ExceptionReporter.IsRunningMono())
					_view.PopulateSysInfoTab();
			}
			finally
			{
				_view.SetProgressCompleteState();
			}
		}

		/// <summary>
		/// Close/cleanup
		/// </summary>
		public void Close()
		{
			_reportGenerator.Dispose();
		}
	}
}