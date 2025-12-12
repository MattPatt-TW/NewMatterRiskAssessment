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

[] HOD Approval takes too long to run... I indentified we have/had 16 separate SQL calls which could be the cause of the slowness, and by looks of it, could be optimised.
