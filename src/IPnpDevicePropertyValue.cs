// Copyright (c) ScaleFS LLC; used with permission
// Licensed under the MIT License

using System.Collections.Generic;

namespace ScaleFS.WindowsPnp;

public interface IPnpDevicePropertyValue
{
     public record ArrayOfValues(List<IPnpDevicePropertyValue> Array) : IPnpDevicePropertyValue;
     public record Boolean(bool Value) : IPnpDevicePropertyValue;
     public record Byte(byte Value) : IPnpDevicePropertyValue;
     public record Guid(System.Guid Value) : IPnpDevicePropertyValue;
     public record ListOfValues(List<IPnpDevicePropertyValue> List) : IPnpDevicePropertyValue;
     public record String(string Value) : IPnpDevicePropertyValue;
     public record UInt16(ushort Value) : IPnpDevicePropertyValue;
     public record UInt32(uint Value) : IPnpDevicePropertyValue;
     public record UnsupportedPropertyDataType(uint/*Windows.Win32.Devices.Properties.DEVPROPTYPE*/ PropertyDataType) : IPnpDevicePropertyValue;
     public record UnsupportedRegistryDataType(uint/*Windows.Win32.System.Registry.REG_VALUE_TYPE*/ RegistryDataType) : IPnpDevicePropertyValue;
}
