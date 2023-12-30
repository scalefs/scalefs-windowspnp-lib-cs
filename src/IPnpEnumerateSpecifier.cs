// Copyright (c) ScaleFS LLC; used with permission
// Licensed under the MIT License

using System;

namespace ScaleFS.WindowsPnp;

public interface IPnpEnumerateSpecifier
{
     public record AllDevices() : IPnpEnumerateSpecifier;
     public record DeviceInterfaceClassGuid(Guid InterfaceClassGuid) : IPnpEnumerateSpecifier;
     public record DeviceSetupClassGuid(Guid SetupClassGuid) : IPnpEnumerateSpecifier;
     public record PnpDeviceInstanceId(String InstanceId, Guid? InterfaceClassGuid) : IPnpEnumerateSpecifier;
     public record PnpEnumeratorId(string EnumeratorId) : IPnpEnumerateSpecifier;
}
