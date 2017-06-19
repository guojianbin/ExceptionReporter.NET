using System;
using System.Reflection;
using ExceptionReporting.Core;
using ExceptionReporting.Views;

// ReSharper disable UnusedMember.Global

namespace ExceptionReporting
{
	/// <summary>
	/// The entry-point (class) to invoking an ExceptionReporter dialog
	/// eg new ExceptionReporter().Show()
	/// </summary>
	public class ExceptionReporter
	{
		private readonly ExceptionReportInfo _reportInfo;
		private IExceptionReportView _view;

		/// <summary>
		/// Initialise the ExceptionReporter
		/// </summary>
		public ExceptionReporter()
		{
			var callingAssembly = Assembly.GetCallingAssembly();

			_reportInfo = new ExceptionReportInfo { AppAssembly = callingAssembly };
		}

		/// <summary>
		/// Public access to configuration
		/// </summary>
		public ExceptionReportInfo Config
		{
			get { return _reportInfo; }
		}

		/// <summary>
		/// Show the ExceptionReport dialog
		/// </summary>
		/// <remarks>The <see cref="ExceptionReporter"/> will analyze the <see cref="Exception"/>s and 
		/// create and show the report dialog.</remarks>
		/// <param name="exceptions">The <see cref="Exception"/>s to show.</param>
		public void Show(params Exception[] exceptions)
		{
			if (exceptions == null) return;   //TODO perhaps show a dialog that says "No exception to show" ?

			try
			{
				_reportInfo.SetExceptions(exceptions);
				_view = new ExceptionReportView(_reportInfo);
				_view.ShowExceptionReport();
			}
			catch (Exception internalException)
			{
				System.Windows.Forms.MessageBox.Show(internalException.Message);
			}
		}

		/// <summary>
		/// Show the ExceptionReport dialog with a custom message instead of the Exception's Message property
		/// </summary>
		/// <param name="customMessage">custom (exception) message</param>
		/// <param name="exceptions">The exception/s to display in the exception report</param>
		public void Show(string customMessage, params Exception[] exceptions)
		{
			_reportInfo.CustomMessage = customMessage;
			Show(exceptions);
		}

		private static readonly bool _isRunningMono = System.Type.GetType("Mono.Runtime") != null;

		/// <returns><c>true</c>, if running mono <c>false</c> otherwise.</returns>
		public static bool IsRunningMono() { return _isRunningMono; }

	}
}

// ReSharper restore UnusedMember.Global
