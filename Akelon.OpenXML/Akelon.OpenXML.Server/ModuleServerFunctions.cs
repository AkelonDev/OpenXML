using System;
using System.Collections.Generic;
using System.Linq;
using Sungero.Core;
using Sungero.CoreEntities;
using System.IO;
//using DocumentFormat.OpenXml.Packaging;
//using DocumentFormat.OpenXml.Wordprocessing;
using Sungero.Domain.Shared;

namespace Akelon.OpenXML.Server
{
  public partial class ModuleFunctions
  {

    /// <summary>
    /// Создать и вернуть Простой документ.
    /// </summary>
    [Remote]
    public Sungero.Docflow.ISimpleDocument CreateSimpleDocument()
    {
      var simpleDcument = Sungero.Docflow.SimpleDocuments.Create();
      simpleDcument.Name = Akelon.OpenXML.Resources.Demo;
      return simpleDcument;
    }
    
    /// <summary>
    /// Получить шаблон для демонстарции решения.
    /// </summary>
    [Remote]
    public Sungero.Docflow.IDocumentTemplate GetTemplate()
    {
      var externalLink = Sungero.Docflow.PublicFunctions.Module.GetExternalLink(typeof(Sungero.Docflow.IDocumentTemplate).GetTypeGuid(), Constants.Module.TemplateExample);
      
      return Sungero.Docflow.DocumentTemplates.GetAll(x => externalLink != null && x.Id == externalLink.EntityId.Value).FirstOrDefault();
    }
    
    /// <summary>
    /// Метод демонстарции заполнения шаблона.
    /// </summary>
    /// <returns>Простой документ с версией для демонстрации.</returns>
    [Remote]
    public Sungero.Docflow.ISimpleDocument ExampleMethod()
    {
      var template = GetTemplate();
      
      if (template == null)
        throw AppliedCodeException.Create(Akelon.OpenXML.Resources.ErrorNoFoundTemplate);
      
      using (var stream = new MemoryStream())
      {
        template.LastVersion.Body.Read().CopyTo(stream);
        
        var result = Akelon.OpenXML.IsolatedFunctions.AkelonOpenXMLWrapper.ExampleWorkWithArea(stream, OpenXML.Resources.Logo.ToString());
        
        var simpleDcument = CreateSimpleDocument();
        simpleDcument.CreateVersionFrom(result, template.AssociatedApplication.Extension);
        
        return simpleDcument;
      }
    }
    
  }
}