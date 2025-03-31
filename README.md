# 📬 Shared Mailbox Permission Tool (PowerShell GUI)

A Windows Forms GUI tool written in PowerShell to **automatically assign FullAccess and Send As permissions** to members of an Active Directory group for a specific Shared Mailbox.

---

## ✨ Features

- 🖤 **Dark theme** interface for comfortable use
- ✅ **Input validation**:
  - Confirms that the AD group exists
  - Confirms that the target mailbox exists and is a Shared Mailbox
- 🧑‍💻 Grants **FullAccess** and **Send As** permissions to all user members
- 📝 **Logs all actions** to the GUI **and an external `.log` file**
- ⚠️ Skips nested groups or non-user objects automatically

---

## 📸 Screenshot

---

## 📦 Requirements

- Run PowerShell **as Administrator**
- Exchange Management Shell (or remote session)
- Active Directory module (`Get-ADGroupMember`)
- Update the domain controller (`$dc`) in the script if needed

---

## 🚀 How to Use

1. Clone this repository or download the `.ps1` script.

2. Open PowerShell **as Administrator**.

3. Run the script:
   ```powershell
   .\SharedMailboxPermissionTool.ps1
