# SSH

## Setup-SSH.ps1
Links `~/.ssh` to the SSH configuration directory of a Git repository.


## SSH key on a TPM virtual smart card
These scripts store an SSH key on a TPM virtual smart card,
so that it can be used without a confirmation prompt on each use.
This is useful for keys that are used by background processes, such as automatic `git fetch` in IDEs.
FIDO2 keys, including the ones stored with Windows Hello (`ecdsa-sk`),
show a Windows Hello prompt for each signature, which is annoying for such use.

The private key is created inside the TPM and cannot be exported.
The PIN of the card is asked once when the key is loaded to the SSH agent.
After that, **any process of the current user can use the key through the agent without prompts**,
until the agent is stopped, the key lifetime expires, or the user logs out.
Therefore, use this key only for purposes where this is acceptable, e.g. Git.

Use Windows Hello or a FIDO2 key with user verification for accessing servers.
For further information, please see the
[SSH page on my website](https://agx.fi/it/ssh.html).


### How it works
- `tpmvscmgr` creates a virtual smart card, whose keys are protected by the TPM.
- The key is accessed from the OpenSSH of Git for Windows using the
  [OpenSC](https://github.com/OpenSC/OpenSC) PKCS#11 module.
- A dedicated `ssh-agent` is run with a fixed socket at `~/.ssh/agent/tpm-vsc.sock`,
  so that it can be configured with `IdentityAgent` in the SSH configuration.
  This way it works also for programs that don't have `SSH_AUTH_SOCK` set, such as IDEs.
  The agent is allowed to load only the OpenSC PKCS#11 module.
- OpenSSH sends the PIN given to `ssh-add -s` to every token that the PKCS#11 module exposes.
  To avoid sending the PIN of the virtual smart card to other cards, such as the PIV applet of a YubiKey,
  OpenSC is used with a dedicated configuration at `%LOCALAPPDATA%\SmartCardSSH\opensc.conf`.
  It allows only the GIDS card driver used by the virtual smart cards, and ignores the other readers.
  The scripts refuse to send the PIN if more than one token is visible.


### Requirements
- Windows 11 with a TPM 2.0
- [Git for Windows](https://gitforwindows.org/)
- OpenSC, which can be installed with `Install-Software.ps1`, or with either winget or Chocolatey:
  ```
  winget install --exact --id OpenSC.OpenSC
  ```
  ```
  choco install opensc
  ```
- Admin rights for creating the virtual smart card

Note that Microsoft has deprecated virtual smart cards in favor of Windows Hello for Business and FIDO2,
but `tpmvscmgr` is still included in Windows 11.


### Setup
1. Create the virtual smart card.
   The script elevates itself.
   ```
   .\New-TPMVirtualSmartCard.ps1
   ```
   Choose a PIN of 8-127 characters.
   Despite the name, the PIN does not have to be numeric.
   It can contain uppercase and lowercase letters, digits and special characters,
   but only printable ASCII characters are allowed, so e.g. `ä` and `ö` cannot be used.
   The minimum length can be changed with `-MinPinLength`.
   The card is created with a random administrator key, so the PIN cannot be reset.
   Use `-Puk` to create the card with a PIN unlock key (PUK), which can be used to unblock the PIN.
   The script prints the name of the reader of the new card.
2. Create the key as a normal user (not elevated).
   Windows asks for the PIN of the card.
   ```
   .\New-SmartCardSSHKey.ps1 -ReaderName "<reader name printed by the previous script>"
   ```
   If `-ReaderName` is omitted, the script asks you to select the reader from a list.
   The public key is saved to `~/.ssh/id_rsa_tpm_vsc_<computer name>.pub`,
   and the settings to `%LOCALAPPDATA%\SmartCardSSH\config.json`.
3. Add the public key to your Git server, e.g. in the
   [SSH key settings of GitHub](https://github.com/settings/keys).
4. Configure SSH to use the agent for the Git server.
   Add the following to `~/.ssh/config` before the other identities:
   ```
   Host github.com
       IdentityAgent ~/.ssh/agent/tpm-vsc.sock
       IdentityFile ~/.ssh/id_rsa_tpm_vsc_<computer name>.pub
       IdentitiesOnly yes
   ```
   `IdentityFile` values accumulate from all matching `Host` and `Match` blocks.
   If other blocks add FIDO2 keys for all hosts, exclude the Git server from them,
   e.g. `Match !host github.com exec "..."`.
   Otherwise SSH falls back to the FIDO2 keys and their prompts whenever the agent is not running.
5. Load the key to the agent. Enter the PIN of the card when asked.
   ```
   .\Start-SmartCardSSHAgent.ps1
   ```
6. Test the connection:
   ```
   ssh -T git@github.com
   ```


### Daily use
Run `Start-SmartCardSSHAgent.ps1` once after each login.
If the agent is not running, SSH connections using it fail without prompts.

| Command                                         | Description                                           |
|-------------------------------------------------|-------------------------------------------------------|
| `.\Start-SmartCardSSHAgent.ps1`                 | Start the agent and load the key, if not done already |
| `.\Start-SmartCardSSHAgent.ps1 -Lifetime 28800` | Load the key for a limited time, in seconds           |
| `.\Start-SmartCardSSHAgent.ps1 -Restart`        | Restart the agent and reload the key                  |
| `.\Start-SmartCardSSHAgent.ps1 -Stop`           | Stop the agent                                        |


### Troubleshooting
- **"OpenSC exposes N tokens, but exactly one is required"**:
  another card is visible to OpenSC.
  Add its reader name to `ignored_readers` in `%LOCALAPPDATA%\SmartCardSSH\opensc.conf`, or remove the card.
  Note that `Start-SmartCardSSHAgent.ps1` re-generates this file on each start,
  ignoring all readers that are present except the one of the virtual smart card.
- **Creating the key fails in `New-SelfSignedCertificate`**:
  if the reader-specific container name is not accepted,
  remove the `Container` line from `New-SmartCardSSHKey.ps1`,
  and select the virtual smart card in the Windows dialog instead.
- **The PIN is blocked**:
  the TPM locks out the card after too many wrong attempts.
  Unblock it with the PUK if the card was created with one.
  Otherwise, destroy the card and create a new one.
- **Listing the cards**: `.\Get-TPMVirtualSmartCard.ps1` shows the name, instance ID and reader of each
  TPM virtual smart card. This does not require admin access.
  `certutil -scinfo` shows all readers and the certificates on the cards.


### Removal
1. Stop the agent: `.\Start-SmartCardSSHAgent.ps1 -Stop`
2. Remove the public key from your Git server and from `~/.ssh`.
3. Remove the certificate from the certificate store of the current user (`certmgr.msc`, Personal).
4. Find the instance ID of the virtual smart card, which is of the format `ROOT\SMARTCARDREADER\000n`:
   ```
   .\Get-TPMVirtualSmartCard.ps1
   ```
   Then destroy the card. The script elevates itself and asks for confirmation.
   ```
   .\Remove-TPMVirtualSmartCard.ps1 -InstanceId "<instance ID>"
   ```
5. Remove `%LOCALAPPDATA%\SmartCardSSH`.


### References
- [Microsoft: Tpmvscmgr](https://learn.microsoft.com/en-us/windows/security/identity-protection/virtual-smart-cards/virtual-smart-card-tpmvscmgr)
- [OpenSC configuration](https://github.com/OpenSC/OpenSC/blob/master/etc/opensc.conf.example.in)
- [OpenSSH PKCS#11 support](https://github.com/openssh/openssh-portable/blob/master/ssh-pkcs11.c)
