// Copyright (c) ScaleFS LLC; used with permission
// Licensed under the MIT License

using System;

namespace ScaleFS.WindowsPnp;

public struct PnpDevicePropertyKey
{
     public Guid fmtid { get; init; }
     public uint pid { get; init; }

     internal Windows.Win32.Devices.Properties.DEVPROPKEY ToDevpropkey()
     {
          return new Windows.Win32.Devices.Properties.DEVPROPKEY() { fmtid = this.fmtid, pid = this.pid };
     }

     internal static PnpDevicePropertyKey From(Windows.Win32.Devices.Properties.DEVPROPKEY value)
     {
          return new PnpDevicePropertyKey() { fmtid = value.fmtid, pid = value.pid };
     }

     internal static PnpDevicePropertyKey From(Windows.Win32.UI.Shell.PropertiesSystem.PROPERTYKEY value)
     {
          return new PnpDevicePropertyKey() { fmtid = value.fmtid, pid = value.pid };
     }
}
