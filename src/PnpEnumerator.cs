// Copyright (c) ScaleFS LLC; used with permission
// Licensed under the MIT License

using ScaleFS.Primitives;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace ScaleFS.WindowsPnp;

public class PnpEnumerator
{
    public interface IEnumerateDevicesError
    {
        public record StringDecodingError() : IEnumerateDevicesError;
        public record StringTerminationDecodingError() : IEnumerateDevicesError;
        public record Win32Error(ushort Win32ErrorCode) : IEnumerateDevicesError;

        public static IEnumerateDevicesError FromIWin32Error(IWin32Error win32Error)
        {
            return win32Error switch
            {
                IWin32Error.Win32Error { Win32ErrorCode: var win32ErrorCode } => new IEnumerateDevicesError.Win32Error(win32ErrorCode),
                _ => throw new Exception("invalid code path")
            };
        }
    }
    //
    public static Result<List<PnpDeviceNodeInfo>, IEnumerateDevicesError> EnumeratePresentDevices()
    {
        List<PnpEnumerateOption> options = new([PnpEnumerateOption.IncludeInstanceProperties, PnpEnumerateOption.IncludeDeviceInterfaceClassProperties, PnpEnumerateOption.IncludeSetupClassProperties, PnpEnumerateOption.IncludeDeviceInterfaceProperties]);
        return PnpEnumerator.EnumeratePresentDevices(new IPnpEnumerateSpecifier.AllDevices(), options);
    }
    //
    public static Result<List<PnpDeviceNodeInfo>, IEnumerateDevicesError> EnumeratePresentDevicesByDeviceInterfaceClass(Guid InterfaceClassGuid)
    {
        List<PnpEnumerateOption> options = new([PnpEnumerateOption.IncludeInstanceProperties, PnpEnumerateOption.IncludeDeviceInterfaceClassProperties, PnpEnumerateOption.IncludeSetupClassProperties, PnpEnumerateOption.IncludeDeviceInterfaceProperties]);
        return PnpEnumerator.EnumeratePresentDevices(new IPnpEnumerateSpecifier.DeviceInterfaceClassGuid(InterfaceClassGuid), options);
    }
    //
    public static Result<List<PnpDeviceNodeInfo>, IEnumerateDevicesError> EnumeratePresentDevicesByDeviceSetupClass(Guid setupClassGuid)
    {
        List<PnpEnumerateOption> options = new([PnpEnumerateOption.IncludeInstanceProperties, PnpEnumerateOption.IncludeDeviceInterfaceClassProperties, PnpEnumerateOption.IncludeSetupClassProperties, PnpEnumerateOption.IncludeDeviceInterfaceProperties]);
        return PnpEnumerator.EnumeratePresentDevices(new IPnpEnumerateSpecifier.DeviceSetupClassGuid(setupClassGuid), options);
    }
    //
    public static Result<List<PnpDeviceNodeInfo>, IEnumerateDevicesError> EnumeratePresentDevicesByPnpEnumeratorId(string enumeratorId)
    {
        List<PnpEnumerateOption> options = new([PnpEnumerateOption.IncludeInstanceProperties, PnpEnumerateOption.IncludeDeviceInterfaceClassProperties, PnpEnumerateOption.IncludeSetupClassProperties, PnpEnumerateOption.IncludeDeviceInterfaceProperties]);
        return PnpEnumerator.EnumeratePresentDevices(new IPnpEnumerateSpecifier.PnpEnumeratorId(enumeratorId), options);
    }
    //
    public static Result<List<PnpDeviceNodeInfo>, IEnumerateDevicesError> EnumeratePresentDevices(IPnpEnumerateSpecifier enumerateSpecifier, List<PnpEnumerateOption> options)
    {
        List<PnpDeviceNodeInfo> result = new();

        // configure our variables based on the enumerate specifier
        //
        string? pnpEnumerator;
        Guid? classGuid;
        Guid? deviceInterfaceClassGuid;
        var flags = Windows.Win32.PInvoke.DIGCF_PRESENT;
        switch (enumerateSpecifier)
        {
            case IPnpEnumerateSpecifier.AllDevices { }:
                {
                    pnpEnumerator = null;
                    classGuid = null;
                    deviceInterfaceClassGuid = null;
                    flags |= Windows.Win32.PInvoke.DIGCF_ALLCLASSES;
                }
                break;
            case IPnpEnumerateSpecifier.DeviceInterfaceClassGuid { InterfaceClassGuid: var interfaceClassGuid }:
                {
                    pnpEnumerator = null;
                    classGuid = interfaceClassGuid;
                    deviceInterfaceClassGuid = interfaceClassGuid;
                    flags |= Windows.Win32.PInvoke.DIGCF_DEVICEINTERFACE;
                }
                break;
            case IPnpEnumerateSpecifier.DeviceSetupClassGuid { SetupClassGuid: var setupClassGuid }:
                {
                    pnpEnumerator = null;
                    classGuid = setupClassGuid;
                    deviceInterfaceClassGuid = null;
                    //flags |= 0;
                }
                break;
            case IPnpEnumerateSpecifier.PnpDeviceInstanceId { InstanceId: var instanceId, InterfaceClassGuid: var optionalInterfaceClassGuid }:
                {
                    pnpEnumerator = instanceId;
                    classGuid = null;
                    deviceInterfaceClassGuid = optionalInterfaceClassGuid;
                    flags |= Windows.Win32.PInvoke.DIGCF_DEVICEINTERFACE | Windows.Win32.PInvoke.DIGCF_ALLCLASSES;
                }
                break;
            case IPnpEnumerateSpecifier.PnpEnumeratorId { EnumeratorId: var enumeratorId }:
                {
                    pnpEnumerator = enumeratorId;
                    classGuid = null;
                    deviceInterfaceClassGuid = null;
                    flags |= Windows.Win32.PInvoke.DIGCF_ALLCLASSES;
                }
                break;
            default:
                throw new Exception("invalid code path");
        }

        // parse options
        //
        var includeInstanceProperties = false;
        var includeDeviceInterfaceClassProperties = false;
        var includeDeviceInterfaceProperties = false;
        var includeSetupClassProperties = false;
        foreach (var option in options)
        {
            switch (option)
            {
                case PnpEnumerateOption.IncludeInstanceProperties:
                    includeInstanceProperties = true;
                    break;
                case PnpEnumerateOption.IncludeDeviceInterfaceClassProperties:
                    includeDeviceInterfaceClassProperties = true;
                    break;
                case PnpEnumerateOption.IncludeDeviceInterfaceProperties:
                    includeDeviceInterfaceProperties = true;
                    break;
                case PnpEnumerateOption.IncludeSetupClassProperties:
                    includeSetupClassProperties = true;
                    break;
                default:
                    throw new Exception("invalid code path");
            }
        }

        // see: https://docs.microsoft.com/en-us/windows/win32/api/setupapi/nf-setupapi-setupdigetclassdevsw
        var handleToDeviceInfoSet = Windows.Win32.PInvoke.SetupDiGetClassDevs(classGuid, pnpEnumerator, Windows.Win32.Foundation.HWND.Null, flags);
        if (handleToDeviceInfoSet is null || handleToDeviceInfoSet.IsInvalid == true)
        {
            var win32ErrorCode = Marshal.GetLastWin32Error();
            return Result.ErrorResult<IEnumerateDevicesError>(IEnumerateDevicesError.FromIWin32Error(IWin32Error.FromInt32(win32ErrorCode)));
        }
        //
        try
        {
            // enumerate all the devices in the device info set
            // NOTE: we use a for loop here, but we intend to exit it early once we find the final device; the upper bound is simply a maximum placeholder; we use this construct so that device_index auto-increments each iteration (even if we call 'continue')
            for (uint deviceIndex = 0; deviceIndex < uint.MaxValue; deviceIndex += 1)
            {
                // capture the device info data for this device; we'll extract several pieces of information from this data set
                //
                var devinfoData = new Windows.Win32.Devices.DeviceAndDriverInstallation.SP_DEVINFO_DATA { cbSize = (uint)Marshal.SizeOf<Windows.Win32.Devices.DeviceAndDriverInstallation.SP_DEVINFO_DATA>() };
                //
                // see: https://learn.microsoft.com/en-us/windows/win32/api/setupapi/nf-setupapi-setupdienumdeviceinfo
                var enumDeviceInfoResult = Windows.Win32.PInvoke.SetupDiEnumDeviceInfo(handleToDeviceInfoSet, deviceIndex, ref devinfoData);
                if (enumDeviceInfoResult == false)
                {
                    var win32ErrorCode = Marshal.GetLastWin32Error();
                    if ((Windows.Win32.Foundation.WIN32_ERROR)win32ErrorCode == Windows.Win32.Foundation.WIN32_ERROR.ERROR_NO_MORE_ITEMS)
                    {
                        // if we are out of items to enumerate, break out of the loop now
                        break;
                    }
                    else
                    {
                        return Result.ErrorResult<IEnumerateDevicesError>(IEnumerateDevicesError.FromIWin32Error(IWin32Error.FromInt32(win32ErrorCode)));
                    }
                }

                // using the device info data, capture the device instance ID for this device
                var getDeviceInstanceIdFromDevinfoDataResult = PnpEnumerator.GetDeviceInstanceIdFromDevinfoData(handleToDeviceInfoSet, devinfoData);
                if (getDeviceInstanceIdFromDevinfoDataResult.IsError)
                {
                    switch (getDeviceInstanceIdFromDevinfoDataResult.Error!)
                    {
                        case IGetDeviceInstanceIdFromDevinfoDataError.StringDecodingError { }:
                            return Result.ErrorResult<IEnumerateDevicesError>(new IEnumerateDevicesError.StringDecodingError());
                        case IGetDeviceInstanceIdFromDevinfoDataError.Win32Error { Win32ErrorCode: var win32ErrorCode }:
                            return Result.ErrorResult<IEnumerateDevicesError>(IEnumerateDevicesError.FromIWin32Error(IWin32Error.FromInt32(win32ErrorCode)));
                        default:
                            throw new Exception("invalid code path");
                    }
                }
                string deviceInstanceId = getDeviceInstanceIdFromDevinfoDataResult.Value!;

                //
                //
                //
                //
                //
                // TODO: capture values for all the other PnpDeviceNodeInfo elements; in the meantime, here are empty fillers
                string? containerId = null;
                Dictionary<PnpDevicePropertyKey, IPnpDevicePropertyValue>? deviceInstanceProperties = null;
                Dictionary<PnpDevicePropertyKey, IPnpDevicePropertyValue>? deviceSetupClassProperties = null;
                string? devicePath = null;
                Dictionary<PnpDevicePropertyKey, IPnpDevicePropertyValue>? deviceInterfaceProperties = null;
                Dictionary<PnpDevicePropertyKey, IPnpDevicePropertyValue>? deviceInterfaceClassProperties = null;
                //
                //
                //
                //
                //

                // add this device node's info to our result list
                var deviceNodeInfo = new PnpDeviceNodeInfo()
                {
                    DeviceInstanceId = deviceInstanceId,
                    ContainerId = containerId,
                    //
                    // device instance properties (optional; these should be available for all devices)
                    DeviceInstanceProperties = deviceInstanceProperties,
                    //
                    // device setup class properties (optional, as they only apply to devnodes with device class guids)
                    DeviceSetupClassProperties = deviceSetupClassProperties,
                    //
                    // interface properties (optional, as they only apply to device interfaces)
                    DevicePath = devicePath,
                    DeviceInterfaceProperties = deviceInterfaceProperties,
                    DeviceInterfaceClassProperties = deviceInterfaceClassProperties,
                };
                result.Add(deviceNodeInfo);
            }
        }
        finally
        {
            handleToDeviceInfoSet.Close();
        }

        return Result.OkResult(result);
    }

    //

    public interface IGetDeviceInstanceIdFromDevinfoDataError
    {
        public record StringDecodingError() : IGetDeviceInstanceIdFromDevinfoDataError;
        public record Win32Error(ushort Win32ErrorCode) : IGetDeviceInstanceIdFromDevinfoDataError;

        public static IGetDeviceInstanceIdFromDevinfoDataError FromIWin32Error(IWin32Error win32Error)
        {
            return win32Error switch
            {
                IWin32Error.Win32Error { Win32ErrorCode: var win32ErrorCode } => new IGetDeviceInstanceIdFromDevinfoDataError.Win32Error(win32ErrorCode),
                _ => throw new Exception("invalid code path")
            };
        }
    }
    //
    private static Result<string, IGetDeviceInstanceIdFromDevinfoDataError> GetDeviceInstanceIdFromDevinfoData(Windows.Win32.SetupDiDestroyDeviceInfoListSafeHandle handleToDeviceInfoSet, Windows.Win32.Devices.DeviceAndDriverInstallation.SP_DEVINFO_DATA devinfoData)
    {
        // get the size of the device instance id, null-terminated, as a count of utf-16 characters; we'll get an error code of ERROR_INSUFFICIENT_BUFFER and the requiredSize prarameter will contain the required size
        uint requiredSize = 0;
        //
        bool getDeviceInstanceIdResult;
        unsafe
        {
            // see: https://learn.microsoft.com/en-us/windows/win32/api/setupapi/nf-setupapi-setupdigetdeviceinstanceidw
            getDeviceInstanceIdResult = Windows.Win32.PInvoke.SetupDiGetDeviceInstanceId(handleToDeviceInfoSet, devinfoData, null, 0, &requiredSize);
        }
        if (getDeviceInstanceIdResult == false)
        {
            var win32ErrorCode = Marshal.GetLastWin32Error();
            if ((Windows.Win32.Foundation.WIN32_ERROR)win32ErrorCode == Windows.Win32.Foundation.WIN32_ERROR.ERROR_INSUFFICIENT_BUFFER)
            {
                // this is the expected error (i.e. the error we intentionally induced); continue
            }
            else
            {
                // otherwise, return the error to our caller
                return Result.ErrorResult<IGetDeviceInstanceIdFromDevinfoDataError>(IGetDeviceInstanceIdFromDevinfoDataError.FromIWin32Error(IWin32Error.FromInt32(win32ErrorCode)));
            }
        }
        else
        {
            Debug.Assert(false, "SetupDiGetDeviceInstanceId returned success when we asked it for the required buffer size; it should always return false in this circumstance (since device ids are null terminated and can therefore never be zero bytes in length)");
            return Result.ErrorResult<IGetDeviceInstanceIdFromDevinfoDataError>(new IGetDeviceInstanceIdFromDevinfoDataError.Win32Error((ushort)Windows.Win32.Foundation.WIN32_ERROR.ERROR_INVALID_DATA));
        }
        //
        if (requiredSize == 0)
        {
            Debug.Assert(false, "Device instance ID has zero bytes (and is required to have at least one byte...the null terminator); aborting.");
            return Result.ErrorResult<IGetDeviceInstanceIdFromDevinfoDataError>(new IGetDeviceInstanceIdFromDevinfoDataError.Win32Error((ushort)Windows.Win32.Foundation.WIN32_ERROR.ERROR_INVALID_DATA));
        }
        //
        // allocate memory for the device instance id via a zeroed char vector; then create a PWSTR instance which uses that vector as its mutable data region
        IntPtr pointerToBuffer = Marshal.AllocHGlobal((int)requiredSize * sizeof(char));
        Windows.Win32.Foundation.PWSTR deviceInstanceIdAsPwstr;
        string deviceInstanceId;
        try
        {
            unsafe
            {
                deviceInstanceIdAsPwstr = new((char*)pointerToBuffer.ToPointer());
            }
            //
            // get the device instance id as a PWSTR 
            unsafe
            {
                // see: https://learn.microsoft.com/en-us/windows/win32/api/setupapi/nf-setupapi-setupdigetdeviceinstanceidw
                getDeviceInstanceIdResult = Windows.Win32.PInvoke.SetupDiGetDeviceInstanceId(handleToDeviceInfoSet, devinfoData, deviceInstanceIdAsPwstr, requiredSize, null);
            }
            if (getDeviceInstanceIdResult == false)
            {
                var win32ErrorCode = Marshal.GetLastWin32Error();
                return Result.ErrorResult<IGetDeviceInstanceIdFromDevinfoDataError>(IGetDeviceInstanceIdFromDevinfoDataError.FromIWin32Error(IWin32Error.FromInt32(win32ErrorCode)));
            }
            //
            var deviceInstanceIdAsNullableString = deviceInstanceIdAsPwstr.ToString();
            if (deviceInstanceIdAsNullableString is not null)
            {
                deviceInstanceId = deviceInstanceIdAsNullableString!;
            }
            else
            {
                Debug.Assert(false, "Device instance ID is an invalid string; aborting.");
                return Result.ErrorResult<IGetDeviceInstanceIdFromDevinfoDataError>(new IGetDeviceInstanceIdFromDevinfoDataError.StringDecodingError());
            }
        }
        finally
        {
            Marshal.FreeHGlobal(pointerToBuffer);
        }

        return Result.OkResult(deviceInstanceId);
    }
}
