/*
Luc Morin (MRN), December 2015

Small script to save labels to a folder created based on a project property 

It adds a new Action to Eplan: SaveLabelsToFolder
This action can be called either from a toolbar button, or from a script

Simply "load" the script in Eplan (don't "run" it). 
*/


/*
The following compiler directive is necessary to enable editing scripts
within Visual Studio.

It requires that the "Conditional compilation symbol" SCRIPTENV be defined 
in the Visual Studio project properties

This is because EPLAN's internal scripting engine already adds "using directives"
when you load the script in EPLAN. Having them twice would cause errors.
*/

#if SCRIPTENV
using Eplan.EplApi.ApplicationFramework;
using Eplan.EplApi.Scripting;
using Eplan.EplApi.Base;
using Eplan.EplApi.Gui;
#endif

/*
On the other hand, some namespaces are not automatically added by EPLAN when
you load a script. Those have to be outside of the previous conditional compiler directive
*/

using System;
using System.IO;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;

public class RegisterScriptMenu
{

    //Our Action declaration
    [DeclareAction("SaveLabelsToFolder")]
    public void Action(string rootPath, string schemeName, string fileName, string tempFile, string projectPropertyLabelingScheme)
    //public void Action()
    {

        /*
        It is not possible in EPLAN scripting to "read" or "get" project properties. This requires API.

        As workaround, we create a labeling scheme based on Table of Contents, which scheme will only output
        a single value to a file, which is the property that we want to use as the "new folder" name.
        
        Said scheme must exist in EPLAN before this script can perform.

        */

        ActionManager oMngr = new Eplan.EplApi.ApplicationFramework.ActionManager();
        Eplan.EplApi.ApplicationFramework.Action oSelSetAction = oMngr.FindAction("selectionset");
        Eplan.EplApi.ApplicationFramework.Action oLabelAction = oMngr.FindAction("label");

        //Verify if actions were found
        if (oSelSetAction == null || oLabelAction == null)
        {
            MessageBox.Show("Could not obtain actions");
            return;
        }

        //ActionCallingContext is used to pass "parameters" into EPLAn Actions
        ActionCallingContext ctx = new ActionCallingContext();

        //Using the "selectionset" Action, get the current project path, used later by labeling action
        ctx.AddParameter("TYPE", "PROJECT");
        bool sRet = oSelSetAction.Execute(ctx);

        string sProject = "";

        if (sRet)
        {
            ctx.GetParameter("PROJECT", ref sProject);
        }
        else
        {
            MessageBox.Show("Could not obtain project path");
            return;
        }

        tempFile = string.Format(@"{0}\{1}", rootPath, tempFile);

        //MessageBox.Show(string.Format("{0}\r\n{1}\r\n{2}\r\n{3}\r\n", rootPath, schemeName, fileName, tempFile));

        //Create a new ActionCallingCOntext to make sure we start with an empty one,
        //and then call the labeling action to obtain the desired project value,
        //which gets written to a temporary file.
        ctx = new ActionCallingContext();
        ctx.AddParameter("PROJECTNAME", sProject);
        ctx.AddParameter("CONFIGSCHEME", projectPropertyLabelingScheme);
        ctx.AddParameter("LANGUAGE", "en_US");
        ctx.AddParameter("DESTINATIONFILE", tempFile);
        oLabelAction.Execute(ctx);

        //Now read the temp file, and read the property that was writen to it by the labeling function.
        //Note: Which property gets written to the temp file is determined in the selected labeling scheme settings

        string[] tempContent = File.ReadAllLines(tempFile);
        File.Delete(tempFile);

        //Verify if file had content
        if (tempContent.Length == 0)
        {
            MessageBox.Show("Property file was empty");
            return;
        }

        string projectProperty = tempContent[0];

        //Verify if value is valid
        if(string.IsNullOrEmpty(projectProperty))
        {
            MessageBox.Show("Could not obtain property value from file");
            return;
        }

        string outputPath = string.Format(@"{0}\{1}\{2}", rootPath, projectProperty, fileName);

        //Create a new ActionCallingCOntext to make sure we start with an empty one,
        //and then call the labeling action to obtain the desired labeling output
        //to the desired folder.
        ctx = new ActionCallingContext();
        ctx.AddParameter("PROJECTNAME", sProject);
        ctx.AddParameter("CONFIGSCHEME", schemeName);
        ctx.AddParameter("LANGUAGE", "en_US");
        ctx.AddParameter("DESTINATIONFILE", outputPath);
        oLabelAction.Execute(ctx);


    }
}
