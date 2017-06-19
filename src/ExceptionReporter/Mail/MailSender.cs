using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
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
		readonly ExceptionReportInfo _config;
		IEmailSendEvent _emailEvent;

		internal MailSender(ExceptionReportInfo reportInfo)
		{
			_config = reportInfo;
		}

		/// <summary>
		/// Send SMTP email, requires following ExceptionReportInfo properties to be set
		/// SmtpPort, SmtpFromAddress, EmailReportAddress
		/// Set SmtpUsername/SmtPassword if SMTP server supports/requires authentication
		/// </summary>
		public async Task<bool> SendSmtpAsync(string exceptionReport, IEmailSendEvent emailEvent)
		{
			_emailEvent = emailEvent;

			var message = new MimeMessage
			{
				Subject = EmailSubject,
				Body = new TextPart("plain") { 
					Text = exceptionReport 
				}
			};
			message.From.Add(new MailboxAddress(_config.SmtpFromAddress, _config.SmtpFromAddress));
			message.To.Add(new MailboxAddress(_config.EmailReportAddress, _config.EmailReportAddress));

			using (var client = new SmtpClient())
			{
				client.Connect(_config.SmtpServer, _config.SmtpPort, 
				               _config.SmtpUseSsl ? SecureSocketOptions.SslOnConnect : SecureSocketOptions.None);

				if (!string.IsNullOrWhiteSpace(_config.SmtpUsername))
				{
					client.Authenticate(_config.SmtpUsername, _config.SmtpPassword);
				}

				bool mailSent = false;
				try
				{
					await client.SendAsync(message);
					_emailEvent.Completed(true);
				}
				catch (Exception ex)
				{
					_emailEvent.Completed(false);
					_emailEvent.ShowError(ex.Message, ex);
				}
				finally
				{
					client.Disconnect(true);
				}

				return mailSent;
			}
		}

		/// <summary>
		/// Send SimpleMAPI email
		/// </summary>
		public void SendMapi(string exceptionReport)
		{
			var mapi = new SimpleMapi();

			mapi.AddRecipient(_config.EmailReportAddress, null, false);

			AttachFiles(new AttachAdapter(mapi));
			mapi.Send(EmailSubject, exceptionReport);
		}

		private void AttachFiles(IAttach attacher)
		{
			var filesToAttach = new List<string>();
			if (_config.FilesToAttach.Length > 0)
			{
				filesToAttach.AddRange(_config.FilesToAttach);
			}
			if (_config.ScreenshotAvailable)
			{
				filesToAttach.Add(ScreenshotTaker.GetImageAsFile(_config.ScreenshotImage));
			}

			var existingFilesToAttach = filesToAttach.Where(File.Exists).ToList();

			foreach (var zf in existingFilesToAttach.Where(f => f.EndsWith(".zip")))
			{
				attacher.Attach(zf);    // attach external zip files separately, admittedly weak detection using just file extension
			}

			var nonzipFilesToAttach = existingFilesToAttach.Where(f => !f.EndsWith(".zip")).ToList();
			if (nonzipFilesToAttach.Any())
			{ // attach all other files (non zip) into our one zip file
				var zipFile = Path.Combine(Path.GetTempPath(), _config.AttachmentFilename);
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
					return _config.MainException.Message.Truncate(100);
				}
				catch (Exception)
				{
					return "Exception Report";
				}
			}
		}
	}
}