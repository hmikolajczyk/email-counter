# 📧 Email Counter App
A desktop application for real-time email tracking.
## 📖 Overview
The **Email Counter App** is a lightweight utility designed to help users monitor the number of emails in desired folder. It also allows generating .csv reports which include the most important details of the emails.
## ✨ Key Features
- **Real-time Monitoring:** Keep track of the number of emails in desired folder.
- **Reports:** Generate reports for further analysis.
## 🛠️ Tech Info
- **Framework:** Avalonia UI (Cross-platform XAML-based GUI)
- **Language:** C# / .NET
- **Architecture:** MVVM (Model-View-ViewModel)
- **Design Tools:** Figma
- **Designed for:** Windows 11
## 🚀 Getting Started
### Prerequisites
- .NET SDK 8.0
- Microsoft Outlook (classic)[^1]
### Installation
1. Clone the repository:
```sh
git clone https://github.com/hmikolajczyk/email-counter.git
```
2. Navigate to the project directory:
```sh
cd email-counter\src\EmailCounter.Gui
```
3. Build and run the application:
```sh
dotnet build
dotnet run
```
## 🚧 Current Status
- [x] - Working GUI
- [x] Working .csv exports
- [ ] **Soon:** Stable v1.0.0 release with standalone binaries
## ⚖️ License
Distributed under the MIT License. This software is provided "as is", without warranty of any kind. See `MIT license` for more information.

---
[^1]: Microsoft Outlook (classic) is required for COM Interop connectivity
