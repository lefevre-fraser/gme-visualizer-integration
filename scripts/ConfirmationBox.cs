using System;
using System.Windows;

class ConfirmationBox
{
	public static int Main(string[] args)
	{
		string label = "Confirmation Box";
		string message = "Do you wish to proceed";
		MessageBoxButton buttonType = MessageBoxButton.YesNoCancel;

		if (args.Length == 3) 
		{
			message = args[0];
			label = args[1];
			try
			{
				buttonType = (MessageBoxButton)Enum.Parse(typeof(MessageBoxButton), args[2], true);
			}
			catch
			{
				buttonType = MessageBoxButton.YesNoCancel;
			}
		}

		MessageBoxResult result = MessageBox.Show(message, String.Format(" {0}", label), buttonType);
		return (int)result;
	}
}