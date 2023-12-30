// Copyright (c) ScaleFS LLC; used with permission
// Licensed under the MIT License

using System.Collections.Generic;

namespace ScaleFS.WindowsPnp;

public struct PnpDeviceNodeInfo
{
     // device instance id (applies to all devices)
     public string DeviceInstanceId { get; init; }
     // NOTE: the container_id is probably unique, but is _not_ guaranteed to be unique; see: https://techcommunity.microsoft.com/t5/microsoft-usb-blog/how-to-generate-a-container-id-for-a-usb-device-part-2/ba-p/270726
     // NOTE: the container_id should be available for virutally all devices, but perhaps _not_ for root hubs (and maybe not for hubs at all...TBD)
     public string? ContainerId { get; init; }
     //
     // device instance properties (optional; these should be available for all devices)
     public Dictionary<PnpDevicePropertyKey, IPnpDevicePropertyValue>? DeviceInstanceProperties { get; init; }
     //
     // device setup class properties (optional) (also, they are available for most devices but not ALL devices)
     public Dictionary<PnpDevicePropertyKey, IPnpDevicePropertyValue>? DeviceSetupClassProperties { get; init; }
     //
     // device path (only applies to device interfaces; will be None otherwise)
     public string? DevicePath { get; init; }
     // interface properties (optional) (also, they only apply to device interfaces; will be None otherwise)
     public Dictionary<PnpDevicePropertyKey, IPnpDevicePropertyValue>? DeviceInterfaceProperties { get; init; }
     // interface class properties (optional) (also, they only apply to device interfaces; will be None otherwise)
     public Dictionary<PnpDevicePropertyKey, IPnpDevicePropertyValue>? DeviceInterfaceClassProperties { get; init; }
}
