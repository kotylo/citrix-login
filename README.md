# What is it?
Autologins into Citrix VM using saved Microsoft credentials. Windows only.

# How to use
- Create `.\Src\.env` file based on `.env.example` and fill your data.
- Create a password in Windows Credentials Store (Start → Manage Windows Credentials) for website `AHK_CredentialsForCitrix`:
<img width="854" height="299" alt="image" src="https://github.com/user-attachments/assets/1b24925d-7767-4576-a362-9e760d460125" />

# Run
- `.\Src\run.ps1` in PowerShell
- Enter password from SMS
- Citrix .ica file should be downloaded into `Downloads` folder for current user
- Citrix VM will open it up and login will be performed
