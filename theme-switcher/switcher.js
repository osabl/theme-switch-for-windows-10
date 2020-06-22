try {
	var shell = WScript.CreateObject('WScript.Shell')
  
  // '0' - dark theme, '1' - light theme
	var currentAppTheme = shell.RegRead('HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize\\AppsUseLightTheme')

	if (currentTheme) {
		// Apps theme
		shell.RegWrite('HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize\\AppsUseLightTheme', '0', 'REG_DWORD')
		// System theme
		shell.RegWrite('HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize\\SystemUsesLightTheme', '0', 'REG_DWORD')
	} else {
		// Apps theme
		shell.RegWrite('HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize\\AppsUseLightTheme', '1', 'REG_DWORD')
		// System theme
		shell.RegWrite('HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize\\SystemUsesLightTheme', '1', 'REG_DWORD')
	}
} catch (err) {
	var shell = WScript.CreateObject('WScript.Shell')
  shell.RegWrite('HKEY_CURRENT_USER\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize\\AppsUseLightTheme', '1', 'REG_DWORD')
  
  WScript.Echo('The key was not found. \nA default key was created with a value of "1" (light theme). \nDon`t worry, with the subsequent launch, the theme should switch correctly.')
}
