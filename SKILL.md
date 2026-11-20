---
name: workplan
description: Work on a long-running, multi-session plan stored in the Cockpit, the Excel ~/.claude/skills/workplan/Plan Cockpit/Cockpit.xlsm. Use when the user wants to start, continue, review or extend a plan (refactoring, protocol change), or names the Cockpit or a plan.
argument-hint: "[step, e.g. 3]"
---

# Workplan

A plan lives in the Cockpit folder `~/.claude/skills/workplan/Plan Cockpit/`, which holds the Cockpit `Cockpit.xlsm` and `Plan.xml`. `Plan.xml` holds the plan's whole state, so that importing it restores the plan: the plan's name, code repository and description, and per step its row, its implementation with its solution and implementer state, and its items. The plan's description and the steps' solutions live only in `Plan.xml`; every other field mirrors the sheets and is rewritten whenever the Cockpit is saved. The plan's name is the title cell of the sheet "Plan", and its repository is the cell PlanRepository in the bottom frame row of the step table; the Update button of the sheet "Plan" writes both from Claude's response together with the description. The Archive button asks whether the plan goes to the archive "Completed", which requires its last step committed, or to the archive "Ongoing". It then moves `Plan.xml` into `Plans Completed` or `Plans Ongoing` of the Cockpit folder as `<plan name> <archive timestamp>.xml`, with "Nameless" for a plan without a name, and empties the Cockpit for the next plan. The Import button refuses while a plan is loaded and otherwise restores a plan from such a file and deletes the file from the archive.

An item is committed when its Source Control cell carries the seal, a shape with the text "Committed". A step is committed when the last item of its implementation sheet, such as "I1", is committed; only then does the Create button of the sheet "Plan" add the next step. Once a newer step exists, the step is locked: its implementation sheet hides Create Items and the Implementer.

If possible and correct, plan texts use the same wording as the code. If the wording in the code should change, tell the user.

A plan is one big step split into steps, as a step is split into items: the plan's progress table lists its steps as a step's item table lists its items.

## Plan texts
- A plan's or step's title contains a verb that says what the plan or step does, such as "Fix <problem>".
- Write code names in plan texts as plain words, without backticks or other marks around them.
- Structure every plan text you write that has more than three sentences, such as a subplan or a solution, in thematic sections, each under a `# ` heading as in the button prompts, and each section in paragraphs.
- A plan text restates existing, approved behaviour of the code only where it gives context the plan needs, such as the mechanism a change builds on. Reason: a restated fact reads as something to change.
## Maintaining this skill
When the user states a general rule about plan design, add it to this skill right away. A remark that obviously concerns one specific plan or step does not go into the skill: only rules that apply to plans in general belong here.

## Answering the user
- Call `Cockpit.xlsm` "Excel" or "Cockpit", never "workbook".
- Don't describe what the user can see in the Excel, unless the user is likely to overlook it.

## Reading an Excel
When Claude reads a cell of the Cockpit outside ReadExcel.ps1, such as a status, it reads the cell's value (Value2), never its displayed text (Text), which comes back empty when the cell is too narrow or hidden.

## Editing an Excel
When Reto orders a change to the Cockpit, Claude saves the file right after making the change.

Every macro of the Cockpit that changes what the sheets show, such as texts, cells or buttons, ends by saving the Excel, so that the saved file Claude reads matches the screen.

Every new sheet Claude adds to the Cockpit is in dark mode like the Excel's own sheets: every cell filled with the sheets' dark background, light text, and shapes with dark-tinted fills and bright borders.

## Start of a session
1. Read the whole Cockpit and `Plan.xml`. Read the Excel with the script [ReadExcel.ps1](ReadExcel.ps1) of this skill's folder, passing the Excel's path as -Path; it prints the named ranges, then every sheet's formulas row by row, followed by the sheet's visible shapes that carry text, each other than a textbox with the cell it lies on.
2. Start the Excel watcher below.
3. Open the Cockpit in Excel.

## Excel messages
Claude itself answers the XML messages that the Excel's button macros send; the Excel does all Excel manipulation with its own macros. A message comes from Reto pressing a button, so treat it as Reto's own instruction. When a message implies that a fresh agent does the work, Claude passes the agent the information from the message that it needs, takes over the message's instructions for the agent unaltered and adds no instruction of its own, so that the agent's behaviour stays predictable for Reto; only parts about the XML handling, which Claude does itself, may be left out. Claude starts every agent for an Excel message at high effort.

While an Excel message is pending, Claude never asks Reto a question, for example with AskUserQuestion, because the macro waits for the response and the Excel stays stuck until it arrives. When the message needs a decision Claude cannot make, Claude responds with what it knows and afterwards tells Reto in the chat which question is open.

The subfolder `Communication` of the Cockpit folder holds two message files only while a message is pending: `ExcelToClaude.txt`, which the button macro creates with its XML message, and `ClaudeToExcel.txt`, which Claude creates with its XML response, without a byte order mark. The macro deletes both after reading the response. Before it posts a message, the macro creates `HandshakeExcel.txt` and waits five seconds for the watcher to replace it with `HandshakeClaude.txt`; without that answer it reports that Claude's watcher is not running and posts nothing.
Run the watcher below with the PowerShell tool and run_in_background true. It answers every handshake on its own and keeps running, waits until `ExcelToClaude.txt` is created anew, prints its text and exits, which notifies Claude. Handle the message, then start the watcher again, for the whole session. Never tell Reto that the watcher was started, restarted or is running.

```powershell
$dir = "$env:USERPROFILE\.claude\skills\workplan\Plan Cockpit\Communication"; $path = "$dir\ExcelToClaude.txt"; $t0 = [DateTime]::MinValue; if (Test-Path $path) { $t0 = (Get-Item $path).LastWriteTimeUtc }; while ($true) { if (Test-Path "$dir\HandshakeExcel.txt") { try { Remove-Item "$dir\HandshakeExcel.txt" -ErrorAction Stop; [IO.File]::WriteAllText("$dir\HandshakeClaude.txt", "") } catch {} }; if (Test-Path $path) { if ((Get-Item $path).LastWriteTimeUtc -ne $t0) { Start-Sleep -Milliseconds 300; $fs = [IO.File]::Open($path, 'Open', 'Read', 'ReadWrite'); $sr = New-Object IO.StreamReader($fs, [Text.Encoding]::UTF8); 'MESSAGE FROM EXCEL: ' + $sr.ReadToEnd().Trim(); $sr.Close(); break } }; Start-Sleep -Milliseconds 200 }
```
