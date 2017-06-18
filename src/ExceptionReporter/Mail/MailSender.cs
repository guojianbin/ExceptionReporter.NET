using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ExceptionReporting.Core;
using ExceptionReporting.Extensions;
using Ionic.Zip;
using MailKit.Net.Smtp;
using MailKit.Security;
using MimeKit;
using Win32Mapi;

namespace ExceptionReporting.Mail
{
	class MailSender
	{
		readonly ExceptionReportInfo _reportInfo;
		IEmailSendEvent _emailEvent;

		internal MailSender(ExceptionReportInfo reportInfo)
		{
			_reportInfo = reportInfo;
		}

		/// <summary>
		/// Send SMTP email, requires following ExceptionReportInfo properties to be set
		/// SmtpPort, SmtpUseSsl, SmtpFromAddress, EmailReportAddress
		/// Set SmtpUsername/Password if SMTP server supports authentication
		/// </summary>
		public void SendSmtp(string exceptionReport, IEmailSendEvent emailEvent)
		{
			_emailEvent = emailEvent;

			var message = new MimeMessage
			{
				Subject = EmailSubject,
				Body = new TextPart("plain") { 
					Text = exceptionReport 
				}
			};
			message.From.Add(new MailboxAddress(_reportInfo.SmtpFromAddress, _reportInfo.SmtpFromAddress));
			message.To.Add(new MailboxAddress(_reportInfo.EmailReportAddress, _reportInfo.EmailReportAddress));

			using (var client = new SmtpClient())
			{
				client.Connect(_reportInfo.SmtpServer, _reportInfo.SmtpPort, _reportInfo.SmtpUseSsl ? SecureSocketOptions.SslOnConnect : SecureSocketOptions.None);

				if (!_reportInfo.SmtpUsername.IsEmpty())
				{
					client.Authenticate(_reportInfo.SmtpUsername, _reportInfo.SmtpPassword);
				}

				try
				{
					client.Send(message);
				} 
				catch (Exception ex) 
				{
					_emailEvent.Completed(false);
					_emailEvent.ShowError(ex.Message, ex);
				}
				client.Disconnect(true);
				_emailEvent.Completed(true);
			}
		}

		/// <summary>
		/// Send SimpleMAPI email
		/// </summary>
		public void SendMapi(string exceptionReport)
		{
			var mapi = new SimpleMapi();

			mapi.AddRecipient(_reportInfo.EmailReportAddress, null, false);

			AttachFiles(new AttachAdapter(mapi));
			mapi.Send(EmailSubject, exceptionReport);
		}

		private void AttachFiles(IAttach attacher)
		{
			var filesToAttach = new List<string>();
			if (_reportInfo.FilesToAttach.Length > 0)
			{
				filesToAttach.AddRange(_reportInfo.FilesToAttach);
			}
			if (_reportInfo.ScreenshotAvailable)
			{
				filesToAttach.Add(ScreenshotTaker.GetImageAsFile(_reportInfo.ScreenshotImage));
			}

			var existingFilesToAttach = filesToAttach.Where(File.Exists).ToList();

			foreach (var zf in existingFilesToAttach.Where(f => f.EndsWith(".zip")))
			{
				attacher.Attach(zf);    // attach external zip files separately, admittedly weak detection using just file extension
			}

			var nonzipFilesToAttach = existingFilesToAttach.Where(f => !f.EndsWith(".zip")).ToList();
			if (nonzipFilesToAttach.Any())
			{ // attach all other files (non zip) into our one zip file
				var zipFile = Path.Combine(Path.GetTempPath(), _reportInfo.AttachmentFilename);
				if (File.Exists(zipFile)) File.Delete(zipFile);

				using (var zip = new ZipFile(zipFile))
				{
					zip.AddFiles(nonzipFilesToAttach, "");
					zip.Save();
				}

				attacher.Attach(zipFile);
			}
		}

		public string EmailSubject
		{
			get
			{
				try
				{
					return _reportInfo.MainException.Message.Truncate(100);
				}
				catch (Exception)
				{
					return "Exception Report";
				}
			}
		}
	}
}