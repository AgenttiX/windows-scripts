function Toggle-Comment {
    <#
    .LINK
        https://superuser.com/a/1730675
    #>
    $file = $psise.CurrentFile
    $text = $file.Editor.SelectedText
    if ($text.StartsWith("<#")) {
        $comment = $text.Substring(2).TrimEnd("#>")
    }
    else
    {
        $comment = "<#" + $text + "#>"
    }
    $file.Editor.InsertText($comment)
}

# https://superuser.com/a/1730675
$psise.CurrentPowerShellTab.AddOnsMenu.Submenus.Add('Toggle Comment', { Toggle-Comment }, 'CTRL+K') | Out-Null
