# SSH

## Setup-SSH.ps1
Links `~/.ssh` to the SSH configuration directory of a Git repository.


## SSH agent of the PowerShell profile
[`Profile.ps1`](../Profile/Profile.ps1) starts the `ssh-agent` of Git for Windows
when a PowerShell session is opened and no `ssh-agent` process is running.
It saves the agent variables to `~/.ssh-agent-info`,
sets `SSH_AUTH_SOCK` and `SSH_AGENT_PID` in each PowerShell session,
and loads the FIDO2 keys to the agent with `Setup-Agent.ps1` of the private scripts repository.
The agent is used for agent forwarding (`ForwardAgent`),
so that e.g. `git pull` can be run on a server accessed with SSH without storing a key on the server.
Since the keys are FIDO2 keys, each use of a forwarded key requires user presence on the local computer,
e.g. a Windows Hello prompt or touching the security key.

Notes:
- `SSH_AUTH_SOCK` is set only in PowerShell sessions.
  Programs started otherwise, such as IDEs and other GUI applications, don't use this agent,
  and instead use the key files directly.
  With FIDO2 keys this causes a prompt for each use,
  which is why the TPM virtual smart card key below uses a separate agent with a fixed socket.
- The agent is started only if no process called `ssh-agent` is running.
  If some other `ssh-agent` is running, such as the one of Windows OpenSSH,
  the agent of Git for Windows is not started, and `~/.ssh-agent-info` may point to an agent that no longer exists.
- The agent runs until logout, or until it is stopped.
  If it is stopped, the next PowerShell session starts a new one.


## SSH key on a TPM virtual smart card
These scripts store an SSH key on a TPM virtual smart card,
so that it can be used without a confirmation prompt on each use.
This is useful for keys that are used by background processes, such as automatic `git fetch` in IDEs.
FIDO2 keys, including the ones stored with Windows Hello (`ecdsa-sk`),
show a Windows Hello prompt for each signature, which is annoying for such use.

The private key is created inside the TPM and cannot be exported.
The PIN of the card is asked once when the key is loaded to the SSH agent.
After that, **any process of the current user can use the key through the agent without prompts**,
until the agent is stopped or the user logs out.
Therefore, use this key only for purposes where this is acceptable, e.g. Git.

Use Windows Hello or a FIDO2 key with user verification for accessing servers.
For further information, please see the
[SSH page on my website](https://agx.fi/it/ssh.html).


### How it works
- `tpmvscmgr` creates a virtual smart card, whose keys are protected by the TPM.
- The key and a self-signed certificate for it are created with the smart card key storage provider of Windows.
  The certificate is stored in the personal certificate store of the current user.
- The Pageant of [PuTTY-CAC](https://github.com/NoMoreFood/putty-cac) loads the key using the certificate.
  It is started with PIN caching enabled and signing confirmation prompts disabled.
  PuTTY-CAC saves these settings to the registry.
- `ssh-pageant` of Git for Windows provides access to Pageant from the OpenSSH of Git for Windows
  using a fixed socket at `~/.ssh/agent/tpm-vsc.sock`,
  which is configured with `IdentityAgent` in the SSH configuration.
  This way it works also for programs that don't have `SSH_AUTH_SOCK` set, such as IDEs,
  and for programs that set `GIT_SSH_COMMAND` to use `ssh`, which overrides `core.sshCommand` of Git.
- `Start-SmartCardSSHAgent.ps1` makes a test signature, so that the PIN is asked when the script is run,
  and not later when a background process uses the key.


### Requirements
- Windows 11 with a TPM 2.0
- [Git for Windows](https://gitforwindows.org/)
- PuTTY-CAC, which can be installed with `Install-Software.ps1`, or with either winget or Chocolatey:
  ```
  winget install --exact --id NoMoreFood.PuTTY-CAC
  ```
  ```
  choco install putty-cac
  ```
  PuTTY-CAC replaces the regular PuTTY, as they use the same installation directory.
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
   If you have already created the key and its certificate, use its thumbprint instead:
   ```
   .\New-SmartCardSSHKey.ps1 -Thumbprint "<certificate thumbprint>"
   ```
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
5. Start the agent. Enter the PIN of the card when asked.
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

| Command                                  | Description                                                   |
|------------------------------------------|---------------------------------------------------------------|
| `.\Start-SmartCardSSHAgent.ps1`          | Start Pageant and ssh-pageant, if not running already         |
| `.\Start-SmartCardSSHAgent.ps1 -Restart` | Restart Pageant and ssh-pageant. Removes all keys from Pageant |
| `.\Start-SmartCardSSHAgent.ps1 -Stop`    | Stop Pageant and ssh-pageant. Removes all keys from Pageant    |


### Troubleshooting
- **"The key was not found in Pageant"**:
  Pageant was already running without the key, and a new instance cannot add keys to it.
  Run the script with `-Restart`, or add the certificate with "Add CAPI Cert" in the tray menu of Pageant.
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


### Why not OpenSC
The first version of these scripts used the [OpenSC](https://github.com/OpenSC/OpenSC) PKCS#11 module
with the `ssh-agent` of Git for Windows. This does not work for the following reasons.

- **ABI incompatibility**:
  the OpenSSH of Git for Windows is built for MSYS2, which is based on Cygwin.
  It cannot use native Windows PKCS#11 modules such as `opensc-pkcs11.dll`,
  and `ssh-keygen -D` and `ssh-add -s` crash with a segmentation fault when calling the module.
  This is most likely caused by the different size of `long`,
  which the PKCS#11 type `CK_ULONG` is based on:
  it is 32 bits in native 64-bit Windows programs (LLP64), but 64 bits in MSYS2 and Cygwin programs (LP64),
  so the structures of the PKCS#11 API don't match.
  See [OpenSC issue #607](https://github.com/OpenSC/OpenSC/issues/607).
- **Windows OpenSSH**: the native Windows build of OpenSSH has the same ABI as OpenSC,
  but its `ssh-agent` service cannot be configured to allow loading PKCS#11 modules.
  See [Win32-OpenSSH issue #2410](https://github.com/PowerShell/Win32-OpenSSH/issues/2410).
  Without an agent, the PIN would be asked for each connection.
- **PIN length**: OpenSC reports the PIN length range of TPM virtual smart cards as 4-15 characters,
  although `tpmvscmgr` allows PINs of up to 127 characters.
  Therefore, OpenSC may refuse PINs longer than 15 characters.
  This has not been tested.
- **PIN sent to all tokens**: OpenSSH sends the PIN given to `ssh-add -s` to every token
  that the PKCS#11 module exposes, which could e.g. use up the PIN retries of the PIV applet of a YubiKey.
  See [`ssh-pkcs11.c`](https://github.com/openssh/openssh-portable/blob/master/ssh-pkcs11.c).

PuTTY-CAC uses the key through the certificate store and the smart card key storage provider of Windows,
so none of these issues apply to it.


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
- [PuTTY-CAC](https://github.com/NoMoreFood/putty-cac)
- [ssh-pageant](https://github.com/cuviper/ssh-pageant)
