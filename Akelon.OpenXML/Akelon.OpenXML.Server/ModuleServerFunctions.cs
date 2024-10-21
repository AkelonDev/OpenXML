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
  public class ModuleFunctions
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
      
      if (externalLink != null)
        return Sungero.Docflow.DocumentTemplates.GetAll(x => x.Id == externalLink.EntityId.Value).FirstOrDefault();
      
      return null;
    }
    
    /// <summary>
    /// Метод демонстарции заполнения шаблона.
    /// </summary>
    /// <returns>Простой документ с версией для демонстрации.</returns>
    [Remote]
    public Sungero.Docflow.ISimpleDocument ExampleMethod()
    {
      // Получить шаблон.
      var template = GetTemplate();
      
      if (template == null)
        throw AppliedCodeException.Create(Akelon.OpenXML.Resources.ErrorNoFoundTemplate);
      
      var body = new System.IO.MemoryStream();
      template.LastVersion.Body.Read().CopyTo(body);
      
      // Выполнить подстановку значений в элементы управления.
      var result = Akelon.OpenXML.IsolatedFunctions.AkelonOpenXMLWrapper.ExampleWorkWithArea(body, OpenXML.Resources.Logo.ToString());
      
      // Создать простой документ для отобрадения выполнненной подстановки.
      var simpleDcument = CreateSimpleDocument();
      simpleDcument.CreateVersionFrom(result, template.AssociatedApplication.Extension);
      
      return simpleDcument;
    }
    
  }
}