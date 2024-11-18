using System;
using System.Collections.Generic;
using System.Linq;
using Sungero.Core;
using Sungero.CoreEntities;
using Sungero.Domain.Initialization;
using System.IO;
using Sungero.Domain.Shared;

namespace Akelon.OpenXML.Server
{
  public partial class ModuleInitializer
  {

    public override void Initializing(Sungero.Domain.ModuleInitializingEventArgs e)
    {
      // Создать шаблон для демонстарции решения.
      CreateTemplate();
    }

    /// <summary>
    /// Создать Шаблон для демонстрации решения из строки base64.
    /// </summary>
    public void CreateTemplate()
    {
      var externalLink = Sungero.Docflow.PublicFunctions.Module.GetExternalLink(typeof(Sungero.Docflow.IDocumentTemplate).GetTypeGuid(), Constants.Module.TemplateExample);
      
      if (externalLink == null)
      {
        var template = Sungero.Docflow.DocumentTemplates.Create();
        template.Name = Akelon.OpenXML.Resources.TemplateDemoSolution;
        var bytes = Convert.FromBase64String(Queries.Module.Template);
        
        using (var stream = new MemoryStream(bytes))
          template.CreateVersionFrom(stream, Akelon.OpenXML.Resources.TemplateExtension);
        
        template.Save();
        
        Sungero.Docflow.PublicFunctions.Module.CreateExternalLink(template, Constants.Module.TemplateExample);
      }
    }
    
  }
}
