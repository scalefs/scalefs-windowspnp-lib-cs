// Copyright (c) ScaleFS LLC; used with permission
// Licensed under the MIT License

using System;
using System.Collections.Generic;

namespace ScaleFS.WindowsPnp;

public struct PnpDeviceNodeInfo
{
    // device instance id (applies to all devices)
    public string DeviceInstanceId { get; init; }
    // NOTE: the BaseContainerId should be available for virutally all devices, but not for bus drivers or special edge cases (e.g. a volume devnode that spans multiple containers); see: https://learn.microsoft.com/en-us/windows-hardware/drivers/install/overview-of-container-ids
    public Guid? BaseContainerId { get; init; }
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
