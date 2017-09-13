namespace CnCrm.WfActivities.HelperCode {
  internal static class Common {
    internal enum AttributeType {
      String = 0,
      Lookup = 1,
      Picklist = 2,
      Money = 3,
      DateTime = 4,
      Integer = 5,
      Decimal = 6,
      Double = 7,
      Uniqueidentifier = 8,
      Boolean = 9,
      Owner = 10,
      Customer = 11
    }

    internal enum PluginMode {
      Synchronous = 0,
      Asynchronous = 1
    }

    internal const string PreImage = "PreImage",
                          UpdateMessage = "Update",
                          CreateMessage = "Create",
                          Target = "Target";
  }

  internal sealed class EntityOptionSetEnum {
    [System.Diagnostics.DebuggerNonUserCode()]
    public static int? GetEnum(Microsoft.Xrm.Sdk.Entity entity, string attributeLogicalName) {
      if (entity.Attributes.ContainsKey(attributeLogicalName)) {
        Microsoft.Xrm.Sdk.OptionSetValue value = entity.GetAttributeValue<Microsoft.Xrm.Sdk.OptionSetValue>(attributeLogicalName);
        if (value != null) {
          return value.Value;
        }
      }
      return null;
    }
  }
}