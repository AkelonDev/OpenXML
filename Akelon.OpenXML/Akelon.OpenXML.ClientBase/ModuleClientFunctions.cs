using System;
using System.Collections.Generic;
using System.Linq;
using Sungero.Core;
using Sungero.CoreEntities;

namespace Akelon.OpenXML.Client
{
  public class ModuleFunctions
  {

    /// <summary>
    /// Демострация заполнения шаблона кастомными свойствами.
    /// </summary>
    [LocalizeFunction("ExampleDialog_ResourceKey", "ExampleDialog_DescriptionResourceKey")]
    public virtual void DialogForTestOnTemplate()
    {
      var simpleDcument = Functions.Module.Remote.ExampleMethod();
      simpleDcument.LastVersion.Open();
      Sungero.Docflow.PublicFunctions.Module.Remote.EvictEntityFromSession(simpleDcument);
    }

  }
}