import powerfactory as pf
app = pf.GetApplication()
app.ClearOutputWindow()
project = app.GetActiveProject()


# Get the active study case
study_case = app.GetActiveStudyCase()
if study_case is None:
    app.PrintError("No active study case found. Please activate a study case.")
    exit()
while True:
  
  
      # Get the specific GrpPage by path relative to the study case
    page_path = r'Desktop\Curve plot.GrpPage'
    page = study_case.GetContents(page_path)[0]
    if not page:
        app.PrintError(f"GrpPage at '{page_path}' not found.")
        exit()
    app.PrintInfo(f"Found GrpPage: {page.GetFullName()}")
    # Show the page to make it visible
    page.Show()
    app.PrintInfo("Showed the GrpPage to ensure visibility for updating.")
    # Check and create first result object if not exists
    res1_name = 'EMT Simulation 1.ElmRes'
    existing_res1 = study_case.GetContents(res1_name)
    if not existing_res1:
        res1 = study_case.CreateObject('ElmRes', 'EMT Simulation 1')
        if res1 is not None:
            app.PrintPlain("Created: " + res1.GetFullName())
        else:
            app.PrintError("Failed to create EMT Simulation 1")
    else:
        res1 = existing_res1[0]
        app.PrintPlain("Already exists: " + res1.GetFullName())
    # Check and create second result object if not exists
    res2_name = 'EMT Simulation 2.ElmRes'
    existing_res2 = study_case.GetContents(res2_name)
    if not existing_res2:
        res2 = study_case.CreateObject('ElmRes', 'EMT Simulation 2')
        if res2 is not None:
            app.PrintPlain("Created: " + res2.GetFullName())
        else:
            app.PrintError("Failed to create EMT Simulation 2")
    else:
        res2 = existing_res2[0]
        app.PrintPlain("Already exists: " + res2.GetFullName())
        
    app.EchoOn()
    # First simulation with initial conditions from EMT Simulation 1
    initial_conditions = app.GetFromStudyCase('ComInc')
    if initial_conditions:
        initial_conditions.iopt_sim = 'ins'
        initial_conditions.p_resvar = res1
        err = initial_conditions.Execute()
        if err == 0:
            app.PrintPlain("Initial conditions for simulation 1 (from EMT Simulation 1) calculated successfully.")
        else:
            app.PrintError("Failed to calculate initial conditions for simulation 1.")
    else:
        app.PrintError("ComInc object not found.")
    sim = app.GetFromStudyCase('ComSim')
    if sim:
        sim.iopt_sim = 'ins'
        sim.p_res = res1
        err = sim.Execute()
        if err == 0:
            app.PrintPlain("EMT simulation 1 executed successfully. Results saved in: " + res1.GetFullName())
        else:
            app.PrintError("Failed to execute EMT simulation 1.")
    else:
        app.PrintError("ComSim object not found.")



    # Second simulation with initial conditions from EMT Simulation 2

    # Find the Library folder (case-insensitive search)
    all_folders = project.GetContents("*.IntPrjfolder")
    library_folder = None
    for f in all_folders:
        if f.GetAttribute("loc_name").lower() == "library":
            library_folder = f
            break
    if not library_folder:
        app.PrintError("Library folder not found. Available folders:")
        for f in all_folders:
            app.PrintPlain(f" - {f.GetAttribute('loc_name')}")
        exit()
    app.PrintInfo(f"Found Library folder: {library_folder.GetFullName()}")
    # Find the Scripts folder within Library (case-insensitive search)
    all_subfolders = library_folder.GetContents("*.IntPrjfolder")
    scripts_folder = None
    for f in all_subfolders:
        if f.GetAttribute("loc_name").lower() == "scripts":
            scripts_folder = f
            break
    if not scripts_folder:
        app.PrintError("Scripts folder not found in Library. Available subfolders:")
        for f in all_subfolders:
            app.PrintPlain(f" - {f.GetAttribute('loc_name')}")
        exit()
    app.PrintInfo(f"Found Scripts folder: {scripts_folder.GetFullName()}")
    # List all DPL scripts in the Scripts folder
    dpl_scripts = scripts_folder.GetContents("*.ComDpl")
    if not dpl_scripts:
        app.PrintError("No DPL scripts found in Scripts folder.")
        exit()
    app.PrintPlain("DPL scripts in Scripts folder:")
    for s in dpl_scripts:
        app.PrintPlain(f" - {s.GetAttribute('loc_name')}")
    # Search for a script whose name contains "DPLscript1" (case-insensitive)
    target_name = "DPLscript1".lower()
    found_script = None
    for s in dpl_scripts:
        if target_name in s.GetAttribute("loc_name").lower():
            found_script = s
            break
    if not found_script:
        app.PrintError(f"Script containing '{target_name}' not found.")
        exit()
    app.PrintInfo(f"Found matching script: {found_script.GetAttribute('loc_name')}")
    # Execute the script
    ret = found_script.Execute()
    if ret == 0:
        app.PrintInfo("DPL script executed successfully.")
    else:
        app.PrintError(f"DPL script execution failed with return code {ret}.")

    #Back to EMT Simulation for Second Scenario
    app.PrintPlain("Python script starting... Doing some work before DPL.")
    initial_conditions = app.GetFromStudyCase('ComInc')
    if initial_conditions:
        initial_conditions.iopt_sim = 'ins'
        initial_conditions.iopt_ins = 1
        initial_conditions.p_resvar = res2
        err = initial_conditions.Execute()
        if err == 0:
            app.PrintPlain("Initial conditions for simulation 2 (from EMT Simulation 2) calculated successfully.")
        else:
            app.PrintError("Failed to calculate initial conditions for simulation 2.")
    else:
        app.PrintError("ComInc object not found.")
    sim = app.GetFromStudyCase('ComSim')
    if sim:
        sim.iopt_sim = 'ins'
        sim.p_res = res2
        err = sim.Execute()
        if err == 0:
            app.PrintPlain("EMT simulation 2 executed successfully. Results saved in: " + res2.GetFullName())
        else:
            app.PrintError("Failed to execute EMT simulation 2.")
    else:
        app.PrintError("ComSim object not found.")
        
    # Search for the prompt DPL Loop script
    prompt_target_name = "DPL Loop".lower()
    prompt_script = None
    for s in dpl_scripts:
        if prompt_target_name in s.GetAttribute("loc_name").lower():
            prompt_script = s
            break
    if not prompt_script:
        app.PrintError(f"Script containing '{prompt_target_name}' not found.")
        break
    app.PrintInfo(f"Found matching prompt script: {prompt_script.GetAttribute('loc_name')}")
    # Execute the prompt script
    ret = prompt_script.Execute()
    if ret != 0:
        app.PrintError(f"Prompt script execution failed with return code {ret}.")
        break
    # Get the user's choice
    error, choice = prompt_script.GetInputParameterInt('i_choice')
    if error != 0:
        app.PrintError("Failed to get choice from prompt.")
        break
    if choice == 0:
        break