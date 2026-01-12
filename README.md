# NewMatterRiskAssessment
All XAML and IronPython code relating to the 'New Matter Risk Assessment' process in P4W

There are a total of 3 screens:
1) Practice-Level - NMRA Setup / Configurator type screen to allow editing, assigning to Case Types, setting score thresholds etc. 
2) Matter-Level - NMRA & File Review on a single matter. Allows user to answer questions and save their answers against a matter.
3) Practice-Level - HOD Approval Screen (for High-Risk matters). Provides HOD with oversight of High-Risk scoring Matters, and allows them to approve continuation on matter.


## Outstanding Issues:

[] Copying/Duplicating an MRA (When HOD approves High-Risk Matter; and if user clicks 'Duplicate')
    - currently, we copy over matter MRA wholly as-is (including duplicating 'header' in 'Usr_MRA_Overview' and the Questions and selected answer and score in the 'Usr_MRA_Detail' table)
    - wondering if in fact we ought to be checking for updated version, and using that version instead? (makes more logical sense)
    - would take a bit of work to re-map, as the ideal scenario is creating less clicks for end-user, so we want to copy over selected answer, so all they need to do is update anything that's different. But if pulling-down updated template, will need to check QuestionText for match to old one, and if match, then lookup answers, and re-select matching one (if found) and getting latest 'score'. If QuestionText doesn't match anything new, then leave 'SelectedAnswerID' to -1 (not selected), and likewise for if there's no match to old AnswerText (in new Answer list)

    This code features on the 'HOD Approval' screen, in addition to the 'Matter-Level' (main Fee Earner) one, as we allow user to 'duplicate MRA'.
    So any updates, we may want to consider putting into its own 'IMPORT' (like with the runSQL() function), and then only using this ONE shared code space, to avoid duplicates.

[] NB: we also have 'HOD Approval' code in the Matter-Level (Fee Earner) screen, so like above, consider moving shared code into dedicated module that can be imported by each screen that wants to make use of.
    - Note: I don't want to be wreckless and chuck all 'reuse' code into ONE 'TWUtils' file, as that becomes awful to manage!
    - So carefully consider what you want to put where and name accordingly.
    - Oh, WIP WriteOff too... we used to allow HOD write-off from main Fee Earner screen, but since I added actual 'write-off' code (that's like 1k lines on its own!), I forced HODs to go into the HOD screen as I didn't want to copy code into the 'Fee Earner' screen... but if I put into separate module, then perhaps we could allow for HODs using the 'Fee Earner' screen for their own matters?

[] HOD Approval takes too long to run... I indentified we have/had 16 separate SQL calls which could be the cause of the slowness, and by looks of it, could be optimised. By CTE?


------------------------------------------------------------------------------------------------------------------------------------------------------------------
# Practice Screen
This is a large screen / lots of tabs of separate info:
- Unlocking Matters (that were locked by incomplete NMRA elapsing)
- Configure Matter Risk Assessment
    - Manage Main Templates
    - Case Type Defaults
    - Editing a Template area
    - Preview area
- Configure File Review
    - Manage Main Templates
    - CaseType/Department Defaults
    - Editing a Template area
    - Preview area
- Manage Global Answers
- HOD Access and Privileges (currently blank... was planning to add Locks and Keys available to users, but didn't want to duplicate 'Locks' XAML!)
- Reporting (currently blank... did want to add some type of reporting here, but never fleshed-out what would be useful to see here)

We recently updated the MRA side to always 'Duplicate' a template, whenever one is selected for 'Editing'.
This copies the main 'template header', 'questions' and 'answers' and takes one into the 'Editing a Template' area for this new item.
One needs to then update that copy as applicable (freely able to delete Q's and A's, as these will have NEVER been used yet), and when happy, click a 'Publish' button.
That 'Publish' button will then overwrite the 'CaseTypeDefaults', looking for the 'old' TypeID and replacing any occurrence with new TypeID.

Whilst we ARE 'linking' to source of a template (TypeID), I'm not sure I have the structure setup optimally.
Louis mentioned having another table to act as main 'joining' table, so that all other tables don't have links to each other (Answers currently linked directly to Questions, and Questions are directly linked to a TemplateType), and we instead have a table of linked IDs.


MRA - Main functions:
- refresh_MRA_Templates()
- DG_MRA_Template_SelectionChanged(s, event)
- btn_SaveMainMRA_Details_Click(s, event)
- DG_MRA_Template_CellEditEnding(s, event)
- AddNew_MRA_Template(s, event)
- btn_Duplicate_MRA_Template(s, event) - - < re-labelled so clearer this is the button clicking function
- Delete_MRA_Template(s, event) 
- Preview_MRA_Template(s, event)
- btn_Edit_MRATemplate(s, event) - NB: using (central) function
- Publish_MRA(s, event) - definitely worth reviewing... lot of SQL calls here
- [NEW] duplicate_MRA_Template(sourceTypeID, newTypeID, newTypeName)
- [NEW] load_MRA_Template_ForEditing(TypeIDtoUse=None, originalItemID=0)
    *** duplicate functions now updated (needs testing) ***

MRA - Score Thresholds funtions:
- btn_SaveScoreThresholds_Click(s, event)
- ST_Low_SliderChanged(s, event)
- ST_Med_SliderChanged(s, event)
- setup_ST_Sliders(s, event)
- addNew_ScoreThreshold(newTypeID)
- duplicate_ScoreThresholds(TypeID_toCopy, newTypeID)
- refresh_MRA_LMH_ScoreThresholds_New(s, event)
- scoreMatrix_setMax(s, event)

MRA - Department Defaults funtions:
* note: we don't have a 'departments defaults' tab... only a 'CaseTypeDefaults'... so in theory we could remove these?
- refresh_MRA_Department_Defaults(s, event)
- MRA_Department_Defaults_SelectionChanged(s, event)
- MRA_Save_Default_For_Department(s, event)
- 

MRA - Case Type Defaults functions:
- add_Missing_CaseTypeDefaults(forWhat='')
- refresh_MRA_CaseType_Defaults(s, event)
- MRA_CaseType_Defaults_SelectionChanged(s, event)
- MRA_Save_Default_For_CaseType(s, event)

    Controls:
    - dg_MRA_CaseTypes_MRATemplate (datagrid on left, showing all Case Types, grouped by department, listing which MRA's apply)
    - dg_MRA_Templates_CTD (datagrid on the right, listing all MRA templates, and links to above/left datagrid to show templates applicable for selected case type)
        ** we currently have issue here because we're not see 'full' list of applicable Templates... code needs updating

MRA - Editing Questions functions:
- refresh_MRA_Questions(s, event)
- MRA_Questions_SelectionChanged(s, event)
- SaveChanges_MRA_Question(s, event)
- dg_EditMRA_AnswersPreview_SelectionChanged(s, event)
- dg_EditMRA_AnswersPreview_CellEditEnding(s, event)
- dg_EditMRA_AnswersPreview_addNew(s, event)
- dg_EditMRA_AnswersPreview_duplicate(s, event)
- dg_EditMRA_AnswersPreview_moveToTop(s, event)
- dg_EditMRA_AnswersPreview_moveUp(s, event)
- dg_EditMRA_AnswersPreview_moveDown(s, event)
- dg_EditMRA_AnswersPreview_moveToBottom(s, event)
- dg_EditMRA_AnswersPreview_deleteSelected(s, event)
- AddNew_MRA_Question(s, event)
- Duplicate_MRA_Question(s, event)
- MoveTop_MRA_Question(s, event)
- MoveUp_MRA_Question(s, event)
- MoveDown_MRA_Question(s, event)
- MoveBottom_MRA_Question(s, event)
- Delete_MRA_Question(s, event)
- 

MRA - Preview area functions:
- refresh_Preview_MRA(s, event)
- MRA_Preview_AutoAdvance(currentDGindex, s, event)
- MRA_Preview_UpdateTotalScore(s, event)
- MRA_Preview_SelectionChanged(s, event)
- populate_MRA_Preview_SelectAnswerCombo(s, event)
- PreviewMRA_BackToOverview(s, event)
- populate_preview_MRA_QGroups(s, event)
- preview_MRA_SaveAnswer(s, event)
- GroupItems_Preview_SelectionChanged(s, event)
- update_EmailComment(s, event)
- 



