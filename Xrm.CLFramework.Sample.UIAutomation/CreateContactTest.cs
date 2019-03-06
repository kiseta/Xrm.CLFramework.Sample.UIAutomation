using System;
using System.Security;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Dynamics365.UIAutomation.Browser;
using Microsoft.Dynamics365.UIAutomation.Api;
using System.Linq;
using OpenQA.Selenium;

namespace Xrm.CLFramework.Sample.UIAutomation
{
	[TestClass]
	public class CreateContactTest
	{
		private readonly SecureString _username = Environment.GetEnvironmentVariable("CRMUser", EnvironmentVariableTarget.User).ToSecureString();
		private readonly SecureString _password = Environment.GetEnvironmentVariable("CRMPassword", EnvironmentVariableTarget.User).ToSecureString();
		private readonly Uri _xrmUri = new Uri(Environment.GetEnvironmentVariable("CRMUrl", EnvironmentVariableTarget.User));

		private readonly BrowserOptions _options = new BrowserOptions
		{
			BrowserType = BrowserType.Chrome,
			PrivateMode = true,
			FireEvents = true
		};

		[TestMethod]
		public void CreateContact()
		{

			//using (var xrmBrowser = new XrmBrowser(_options))
			using (var xrmBrowser = new Browser(_options))

			{
				xrmBrowser.LoginPage.Login(_xrmUri, _username, _password);
				xrmBrowser.GuidedHelp.CloseGuidedHelp();
				xrmBrowser.Navigation.OpenSubArea("Sales", "Contacts");

				xrmBrowser.ThinkTime(2000);
				xrmBrowser.Grid.SwitchView("Active Contacts");

				xrmBrowser.ThinkTime(1000);
				xrmBrowser.CommandBar.ClickCommand("New");

				xrmBrowser.ThinkTime(5000);

				var fields = new List<Field>
					{
							new Field() {Id = "firstname", Value = "Wael"},
							new Field() {Id = "lastname", Value = "Test"}
					};
				xrmBrowser.Entity.SetValue(new CompositeControl() { Id = "fullname", Fields = fields });
				xrmBrowser.Entity.SetValue("emailaddress1", "testfn.testln@gmail.com");
				xrmBrowser.Entity.SetValue("mobilephone", "555-555-5555");
				xrmBrowser.Entity.SetValue("birthdate", DateTime.Parse("11/1/1980"));
				xrmBrowser.Entity.SetValue(new OptionSet { Name = "preferredcontactmethodcode", Value = "Email" });
				xrmBrowser.CommandBar.ClickCommand("Save");
				xrmBrowser.ThinkTime(5000);
				string screenShot = string.Format("{0}\\CreateNewContact.png", TestContext.TestResultsDirectory);
				xrmBrowser.TakeWindowScreenShot(screenShot, ScreenshotImageFormat.Png);
				TestContext.AddResultFile(screenShot);
			}

		}
	}
}
