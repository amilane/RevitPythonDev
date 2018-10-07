#region Namespaces
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Microsoft.Scripting.Hosting;

using IronPython.Hosting;
using Microsoft.Scripting;
using rvtPythonInterop.Properties;

#endregion

namespace rvtPythonInterop
{
  [Transaction(TransactionMode.Manual)]
  public class Command : IExternalCommand
  {
    public Result Execute(
      ExternalCommandData commandData,
      ref string message,
      ElementSet elements)
    {
      UIApplication uiapp = commandData.Application;
      UIDocument uidoc = uiapp.ActiveUIDocument;
      Application app = uiapp.Application;
      Document doc = uidoc.Document;

      var collector = new FilteredElementCollector(doc)
        .OfClass(typeof(Wall))
        .ToElements();

      // Create python engine
      var opts = new Dictionary<string, object>();
      if (Debugger.IsAttached)
        opts["Debug"] = true;
      ScriptEngine engine = Python.CreateEngine(opts);

      // Get script source (path is absolute to the current assembly)
      // .py scripts must be copied to output directory
      string scriptPath = @"D:\#Coding\rvtTest2\rvtPythonInterop\Scripts\script.py";
      var source = GetScriptSource(scriptPath, engine);

      // Create scope for Python engine execution
      ScriptScope scope = engine.CreateScope();
      scope.SetVariable("doc", doc);
      scope.SetVariable("uidoc", uidoc);
      scope.SetVariable("uiapp", uiapp);
      scope.SetVariable("walls", collector);

      dynamic res = source.Execute(scope);

      return Result.Succeeded;
    }

    public ScriptSource GetScriptSource(string scriptFullPathName, ScriptEngine engine)
    {
      if (System.Diagnostics.Debugger.IsAttached)
      {
        return engine.CreateScriptSourceFromFile(scriptFullPathName);
      }

      string scriptName = Assembly.GetExecutingAssembly().GetName().Name + ".Scripts." + Path.GetFileName(scriptFullPathName);
      Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(scriptName);

      string script = new StreamReader(stream).ReadToEnd();

      return engine.CreateScriptSourceFromString(script);
    }
  }
}
