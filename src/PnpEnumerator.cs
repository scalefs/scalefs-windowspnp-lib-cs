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

                // for all devices: capture the base container id of the device
                //
                // NOTE: we could probably also get this data using the modern setup API by retrieving the device instance property "DEVPKEY_Device_BaseContainerId"...which might be preferable to using the legacy device registry property value mechanism; note that its type is GUID instead of String
                // NOTE: SPDRP_BASE_CONTAINERID is not listed as an allowed property at https://learn.microsoft.com/en-us/windows/win32/api/setupapi/nf-setupapi-setupdigetdeviceregistrypropertyw -- this may an additional reason to look at transitioning this call to the modern setup API
                var getDeviceRegistryPropertyValueResult = PnpEnumerator.GetDeviceRegistryPropertyValue(handleToDeviceInfoSet, devinfoData, Windows.Win32.PInvoke.SPDRP_BASE_CONTAINERID);
                if (getDeviceRegistryPropertyValueResult.IsError)
                {
                    switch (getDeviceRegistryPropertyValueResult.Error!)
                    {
                        case IGetDevicePropertyValueError.StringDecodingError:
                            Debug.Assert(false, "BUG: Win32 setupapi's string was not properly terminated with a null terminator.");
                            return Result.ErrorResult<IEnumerateDevicesError>(new IEnumerateDevicesError.StringDecodingError());
                        case IGetDevicePropertyValueError.StringListTerminationError:
                            Debug.Assert(false, "BUG: Win32 setupapi's list of strings was not properly terminated with an extra null terminator.");
                            return Result.ErrorResult<IEnumerateDevicesError>(new IEnumerateDevicesError.StringTerminationDecodingError());
                        case IGetDevicePropertyValueError.StringTerminationError:
                            Debug.Assert(false, "BUG: Win32 setupapi's string (or final string in a list of strings) was not properly terminated with a null terminator.");
                            return Result.ErrorResult<IEnumerateDevicesError>(new IEnumerateDevicesError.StringTerminationDecodingError());
                        case IGetDevicePropertyValueError.Win32Error { Win32ErrorCode: var win32ErrorCode }:
                            return Result.ErrorResult<IEnumerateDevicesError>(new IEnumerateDevicesError.Win32Error(win32ErrorCode));
                        default:
                            throw new Exception("invalid code path");
                    }
                }
                //
                Guid? baseContainerId;
                switch (getDeviceRegistryPropertyValueResult.Value!)
                {
                    case IPnpDevicePropertyValue.String { Value: var value }:
                        if (Guid.TryParse(value, out Guid outputGuid) == true)
                        {
                            if (outputGuid != Guid.Empty)
                            {
                                baseContainerId = outputGuid;
                            }
                            else
                            {
                                // a zeroed GUID value indicates that there is no container
                                // see: https://learn.microsoft.com/en-us/windows-hardware/drivers/install/overview-of-container-ids
                                baseContainerId = null;
                            }
                        } else { 
                            Debug.Assert(false, "GetDeviceRegistryPropertyValue returned an invalid (non-Guid) string value for SPDRP_BASE_CONTAINERID");
                            return Result.ErrorResult<IEnumerateDevicesError>(new IEnumerateDevicesError.Win32Error((ushort)Windows.Win32.Foundation.WIN32_ERROR.ERROR_INVALID_DATA));
                        }
                        break;
                    default:
                        Debug.Assert(false, "GetDeviceRegistryPropertyValue returned a non-string value for SPDRP_BASE_CONTAINERID");
                        return Result.ErrorResult<IEnumerateDevicesError>(new IEnumerateDevicesError.Win32Error((ushort)Windows.Win32.Foundation.WIN32_ERROR.ERROR_INVALID_DATA));
                };

                //
                //
                //
                //
                //
                // TODO: capture values for all the other PnpDeviceNodeInfo elements; in the meantime, here are empty fillers
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
                    BaseContainerId = baseContainerId,
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

    //

    internal interface IGetDevicePropertyValueError
    {
        public record StringDecodingError() : IGetDevicePropertyValueError;
        public record StringListTerminationError() : IGetDevicePropertyValueError;
        public record StringTerminationError() : IGetDevicePropertyValueError;
        public record Win32Error(ushort Win32ErrorCode) : IGetDevicePropertyValueError;

        public static IGetDevicePropertyValueError FromIWin32Error(IWin32Error win32Error)
        {
            return win32Error switch
            {
                IWin32Error.Win32Error { Win32ErrorCode: var win32ErrorCode } => new IGetDevicePropertyValueError.Win32Error(win32ErrorCode),
                _ => throw new Exception("invalid code path")
            };
        }
    }

    private static Result<Unit, IGetDevicePropertyValueError> CheckSetupDiGetDeviceXXXPropertyRequiredSizeResult(bool setupDiGetDeviceXXXPropertyResult, uint requiredSize)
    {
        if (setupDiGetDeviceXXXPropertyResult == false)
        {
            var win32ErrorCode = Marshal.GetLastWin32Error();
            if ((Windows.Win32.Foundation.WIN32_ERROR)win32ErrorCode == Windows.Win32.Foundation.WIN32_ERROR.ERROR_INSUFFICIENT_BUFFER)
            {
                // this is the expected error condition; we'll resize our buffer to match requiredSize
            }
            else
            {
                return Result.ErrorResult<IGetDevicePropertyValueError>(IGetDevicePropertyValueError.FromIWin32Error(IWin32Error.FromInt32(win32ErrorCode)));
            }
        }
        else
        {
            // we don't expect the operation to succeed with a null buffer and zero-length buffer size.
            Debug.Assert(false, "SetupDiGetDeviceXXXPropertyResultW succeeded, even though we passed it no buffer.");

            // return an error if requiredSize is non-zero; otherwise, continue with the understanding that the property has a size of zero
            if (requiredSize > 0)
            {
                return Result.ErrorResult<IGetDevicePropertyValueError>(new IGetDevicePropertyValueError.Win32Error((ushort)Windows.Win32.Foundation.WIN32_ERROR.ERROR_INVALID_DATA));
            }
        }

        return Result.OkResult();
    }

    //

    private static Result<IPnpDevicePropertyValue, IGetDevicePropertyValueError> GetDeviceRegistryPropertyValue(Windows.Win32.SetupDiDestroyDeviceInfoListSafeHandle handleToDeviceInfoSet, Windows.Win32.Devices.DeviceAndDriverInstallation.SP_DEVINFO_DATA deviceInfoData, uint propertyKey)
    {
        // get the type and size of the device registry property
        // see: https://learn.microsoft.com/en-us/windows/win32/api/setupapi/nf-setupapi-setupdigetdeviceregistrypropertyw
        uint propertyRegistryDataTypeAsUint = 0;
        uint requiredSize = 0;
        bool getDeviceRegistryPropertyResult;
        unsafe
        {
            // NOTE: CsWin32's overload of SetupDiGetDeviceRegistryPropertyW which passes back the required size in the initial pass is a special overload which omits the 'PropertyBufferSize' argument
            getDeviceRegistryPropertyResult = Windows.Win32.PInvoke.SetupDiGetDeviceRegistryProperty(handleToDeviceInfoSet, deviceInfoData, propertyKey, &propertyRegistryDataTypeAsUint, null, &requiredSize);
        }
        var checkRequiredSizeResult = PnpEnumerator.CheckSetupDiGetDeviceXXXPropertyRequiredSizeResult(getDeviceRegistryPropertyResult, requiredSize);
        if (checkRequiredSizeResult.IsError) { return Result.ErrorResult(checkRequiredSizeResult.Error!); }

        // retrieve the property
        Span<byte> propertyBufferAsSpan = stackalloc byte[(int)requiredSize];
        propertyBufferAsSpan.Clear(); // .Fill(0);
        unsafe
        {
            getDeviceRegistryPropertyResult = Windows.Win32.PInvoke.SetupDiGetDeviceRegistryProperty(handleToDeviceInfoSet, deviceInfoData, propertyKey, &propertyRegistryDataTypeAsUint, propertyBufferAsSpan, null);
        }
        if (getDeviceRegistryPropertyResult == false)
        {
            var win32ErrorCode = Marshal.GetLastWin32Error();
            return Result.ErrorResult(IGetDevicePropertyValueError.FromIWin32Error(IWin32Error.FromInt32(win32ErrorCode)));
        }

        // map the registry property type to the modern "Windows Vista" device property data type
        var propertyRegistryDataType = (Windows.Win32.System.Registry.REG_VALUE_TYPE)propertyRegistryDataTypeAsUint;
        Windows.Win32.Devices.Properties.DEVPROPTYPE propertyType;
        switch (propertyRegistryDataType)
        {
            case Windows.Win32.System.Registry.REG_VALUE_TYPE.REG_DWORD:
                propertyType = Windows.Win32.Devices.Properties.DEVPROPTYPE.DEVPROP_TYPE_UINT32;
                break;
            case Windows.Win32.System.Registry.REG_VALUE_TYPE.REG_MULTI_SZ:
                // NOTE: in CsWin32, DEVPROP_TYPEMOD_LIST is a uint (as it's a modifier, rather than a DEVPROPTYPE)
                propertyType = Windows.Win32.Devices.Properties.DEVPROPTYPE.DEVPROP_TYPE_STRING | (Windows.Win32.Devices.Properties.DEVPROPTYPE)Windows.Win32.PInvoke.DEVPROP_TYPEMOD_LIST;
                break;
            case Windows.Win32.System.Registry.REG_VALUE_TYPE.REG_SZ:
                propertyType = Windows.Win32.Devices.Properties.DEVPROPTYPE.DEVPROP_TYPE_STRING;
                break;
            default:
                Debug.Assert(false, "Unknown registry data type; consider supporting this data type");
                return Result.OkResult<IPnpDevicePropertyValue>(new IPnpDevicePropertyValue.UnsupportedRegistryDataType((uint)propertyRegistryDataType));
        };

        // convert the property buffer into a property value
        var propertyValueOrErrorResult = PnpEnumerator.ConvertPropertyBufferIntoDevicePropertyValue(propertyBufferAsSpan.ToArray(), propertyType);
        if (propertyValueOrErrorResult.IsError)
        {
            return Result.ErrorResult(propertyValueOrErrorResult.Error!);
        }
        //
        // NOTE: as we are sharing the ConvertPropertyBufferIntoDevicePropertyValue function (i.e. using the DevicePropertyValue's function for DeviceRegistryPropertyValues) with the non-registry proeprty data types, we need to remap the "unsupported data type" back to the actual registry property data type
        var propertyValue = propertyValueOrErrorResult.Value!;
        var returnPropertyValue = propertyValue switch
        {
            IPnpDevicePropertyValue.UnsupportedPropertyDataType { PropertyDataType: _ } => new IPnpDevicePropertyValue.UnsupportedRegistryDataType(propertyRegistryDataTypeAsUint),
            _ => propertyValue, // pass through all other values
        };

        return Result.OkResult(returnPropertyValue);
    }

    //

    private static Result<IPnpDevicePropertyValue, IGetDevicePropertyValueError> ConvertPropertyBufferIntoDevicePropertyValue(byte[] propertyBuffer, Windows.Win32.Devices.Properties.DEVPROPTYPE propertyType)
    {
        var propertyBufferSize = propertyBuffer.Length;

        var propertyValueIsArray = false;
        var propertyValueIsList = false;
        //
        var propertyTypeMask = PnpEnumerator.CalculateMaskToFitValue(Windows.Win32.PInvoke.MAX_DEVPROP_TYPE);
        var propertyTypeModsMask = propertyTypeMask ^ PnpEnumerator.CalculateMaskToFitValue(Windows.Win32.PInvoke.MAX_DEVPROP_TYPEMOD);
        //
        // extract the property type modifier (if any) from the passed-in property type
        var propertyTypeMod = (uint)propertyType & propertyTypeModsMask;
        //
        // strip any specified mod from the property type value
        var propertyTypeWithoutMods = (Windows.Win32.Devices.Properties.DEVPROPTYPE)((uint)propertyType & propertyTypeMask);

        switch (propertyTypeMod)
        {
            case 0:
                // no mods
                break;
            case Windows.Win32.PInvoke.DEVPROP_TYPEMOD_ARRAY:
                {
                    // see: https://docs.microsoft.com/en-us/windows-hardware/drivers/install/devprop-typemod-array
                    switch ((uint)propertyTypeWithoutMods)
                    {
                        case (uint)Windows.Win32.Devices.Properties.DEVPROPTYPE.DEVPROP_TYPE_BYTE:
                        case (uint)Windows.Win32.Devices.Properties.DEVPROPTYPE.DEVPROP_TYPE_BOOLEAN:
                        case (uint)Windows.Win32.Devices.Properties.DEVPROPTYPE.DEVPROP_TYPE_GUID:
                        case (uint)Windows.Win32.Devices.Properties.DEVPROPTYPE.DEVPROP_TYPE_UINT16:
                        case (uint)Windows.Win32.Devices.Properties.DEVPROPTYPE.DEVPROP_TYPE_UINT32:
                            // these fixed-length value types are allowed
                            propertyValueIsArray = true;
                            break;
                        default:
                            // no other types are allowed
                            Debug.Assert(false, "Device type mod should only be applied to string types; do we have a new string type to handle?");
                            return Result.OkResult<IPnpDevicePropertyValue>(new IPnpDevicePropertyValue.UnsupportedPropertyDataType((uint)propertyType));
                    }
                }
                break;
            case Windows.Win32.PInvoke.DEVPROP_TYPEMOD_LIST:
                {
                    // see: https://docs.microsoft.com/en-us/windows-hardware/drivers/install/devprop-typemod-list
                    switch ((uint)propertyTypeWithoutMods)
                    {
                        case (uint)Windows.Win32.Devices.Properties.DEVPROPTYPE.DEVPROP_TYPE_STRING:
                        case (uint)Windows.Win32.Devices.Properties.DEVPROPTYPE.DEVPROP_TYPE_SECURITY_DESCRIPTOR_STRING:
                            // these string types are allowed
                            propertyValueIsList = true;
                            break;
                        default:
                            // no other types are allowed
                            Debug.Assert(false, "Device type mod should only be applied to string types; do we have a new string type to handle?");
                            return Result.OkResult<IPnpDevicePropertyValue>(new IPnpDevicePropertyValue.UnsupportedPropertyDataType((uint)propertyType));
                    }
                }
                break;
            default:
                // if there are any property type mods which we don't handle, return UnsupportedType (and assert during debug, so we know there are new mods to handle)
                Debug.Assert(false, "Unhandled device type mod");
                return Result.OkResult<IPnpDevicePropertyValue>(new IPnpDevicePropertyValue.UnsupportedPropertyDataType((uint)propertyType));
        }

        // NOTE: all of the fixed-sized value types follow the same pattern (i.e. is or is not array, fixed-size data copied from a byte buffer, wrapped and returned as an IPnpDevicePropertyValue) so we wrap up the common functionality in a named closure (and let it use a type-specific named closure to do the value parsing)
        Func<int, bool, Func<byte[], IPnpDevicePropertyValue>, Result <IPnpDevicePropertyValue, IGetDevicePropertyValueError>> createReturnValueForFixedSizePropertyFunc = delegate (int fixedSizeOfValueType, bool propertyValueIsArray, Func<byte[], IPnpDevicePropertyValue> bufferToPnpDevicePropertyValueFunc)
        {
            if (propertyValueIsArray == true)
            {
                if (propertyBufferSize % fixedSizeOfValueType != 0)
                {
                    Debug.Assert(false, "Invalid property value size");
                    return Result.ErrorResult<IGetDevicePropertyValueError>(new IGetDevicePropertyValueError.Win32Error((int)Windows.Win32.Foundation.WIN32_ERROR.ERROR_INVALID_DATA));
                }

                List<IPnpDevicePropertyValue> arrayOfPropertyValues = new();
                for (var index = 0; index < propertyBufferSize; index += fixedSizeOfValueType)
                {
                    byte[] propertyBufferSlice = new byte[fixedSizeOfValueType];
                    Array.Copy(propertyBuffer, index, propertyBufferSlice, 0, fixedSizeOfValueType);
                    //
                    arrayOfPropertyValues.Add(bufferToPnpDevicePropertyValueFunc(propertyBufferSlice));
                }

                return Result.OkResult<IPnpDevicePropertyValue>(new IPnpDevicePropertyValue.ArrayOfValues(arrayOfPropertyValues));
            }
            else
            {
                if (propertyBufferSize != fixedSizeOfValueType)
                {
                    Debug.Assert(false, "Invalid property value size");
                    return Result.ErrorResult<IGetDevicePropertyValueError>(new IGetDevicePropertyValueError.Win32Error((int)Windows.Win32.Foundation.WIN32_ERROR.ERROR_INVALID_DATA));
                }

                return Result.OkResult(bufferToPnpDevicePropertyValueFunc(propertyBuffer));
            }
        };

        switch ((uint)propertyTypeWithoutMods)
        {
            case (uint)Windows.Win32.Devices.Properties.DEVPROPTYPE.DEVPROP_TYPE_BYTE:
                {
                    const int fixedSizeOfValueType = 1;

                    Func<byte[], IPnpDevicePropertyValue> bufferToPnpDevicePropertyValueFunc = delegate (byte[] buffer)
                    {
                        // convert the byte array to a uint8
                        byte value = buffer[0];
                        return new IPnpDevicePropertyValue.Byte(value);
                    };

                    return createReturnValueForFixedSizePropertyFunc(fixedSizeOfValueType, propertyValueIsArray, bufferToPnpDevicePropertyValueFunc);
                }
            case (uint)Windows.Win32.Devices.Properties.DEVPROPTYPE.DEVPROP_TYPE_BOOLEAN:
                {
                    const int fixedSizeOfValueType = 1;

                    Func<byte[], IPnpDevicePropertyValue> bufferToPnpDevicePropertyValueFunc = delegate (byte[] buffer)
                    {
                        // convert the byte array to a bool
                        bool value = buffer[0] != 0;
                        return new IPnpDevicePropertyValue.Boolean(value);
                    };

                    return createReturnValueForFixedSizePropertyFunc(fixedSizeOfValueType, propertyValueIsArray, bufferToPnpDevicePropertyValueFunc);
                }
            case (uint)Windows.Win32.Devices.Properties.DEVPROPTYPE.DEVPROP_TYPE_GUID:
                {
                    const int fixedSizeOfValueType = 16;

                    Func<byte[], IPnpDevicePropertyValue> bufferToPnpDevicePropertyValueFunc = delegate (byte[] buffer)
                    {
                        // convert the byte array to a guid (using native endian)
                        Guid value = new(
                                  a: BitConverter.ToUInt32(buffer, 0),
                                  b: BitConverter.ToUInt16(buffer, 4),
                                  c: BitConverter.ToUInt16(buffer, 6),
                                  d: buffer[8],
                                  e: buffer[9],
                                  f: buffer[10],
                                  g: buffer[11],
                                  h: buffer[12],
                                  i: buffer[13],
                                  j: buffer[14],
                                  k: buffer[15]
                             );
                        return new IPnpDevicePropertyValue.Guid(value);
                    };

                    return createReturnValueForFixedSizePropertyFunc(fixedSizeOfValueType, propertyValueIsArray, bufferToPnpDevicePropertyValueFunc);
                }
            case (uint)Windows.Win32.Devices.Properties.DEVPROPTYPE.DEVPROP_TYPE_STRING:
                {
                    if (propertyBufferSize % 2 != 0)
                    {
                        Debug.Assert(false, "Invalid property value size");
                        return Result.ErrorResult<IGetDevicePropertyValueError>(new IGetDevicePropertyValueError.Win32Error((int)Windows.Win32.Foundation.WIN32_ERROR.ERROR_INVALID_DATA));
                    }

                    var propertyValueAsUtf16Chars = new List<char>(System.Text.Encoding.Unicode.GetChars(propertyBuffer));
                    if (propertyValueAsUtf16Chars.Count == 0)
                    {
                        Debug.Assert(false, "Invalid property value size; strings and string lists must be null-terminated");
                        return Result.ErrorResult<IGetDevicePropertyValueError>(new IGetDevicePropertyValueError.Win32Error((int)Windows.Win32.Foundation.WIN32_ERROR.ERROR_INVALID_DATA));
                    }

                    if (propertyValueIsList == true)
                    {
                        // NOTE: this list is effectively a REG_MULTI_SZ; it has a final null terminator which terminates the list (and should not be interpreted as an empty string)
                        if (propertyValueAsUtf16Chars[propertyValueAsUtf16Chars.Count - 1] == '\0')
                        {
                            // this is the correct terminator value; proceed
                        }
                        else
                        {
                            return Result.ErrorResult<IGetDevicePropertyValueError>(new IGetDevicePropertyValueError.StringListTerminationError());
                        }
                        propertyValueAsUtf16Chars.RemoveAt(propertyValueAsUtf16Chars.Count - 1);

                        // NOTE: if the list is not an empty list, the final string must be null terminated
                        if (propertyValueAsUtf16Chars.Count > 0)
                        {
                            // NOTE: we are not removing the final character at this time; we're just verifying that the last string in the list is indeed null-terminated
                            if (propertyValueAsUtf16Chars[propertyValueAsUtf16Chars.Count - 1] == '\0')
                            {
                                // this is the correct terminator value; proceed
                            }
                            else
                            {
                                return Result.ErrorResult<IGetDevicePropertyValueError>(new IGetDevicePropertyValueError.StringTerminationError());
                            }
                        }

                        // list of null-terminated strings, separated by their null terminators
                        List<IPnpDevicePropertyValue> listOfStrings = new();

                        // NOTE: this function relies on the fact that the last string should also be null-terminated
                        List<char> currentStringAsUtf16Chars = new();
                        foreach (var utf16Char in propertyValueAsUtf16Chars)
                        {
                            currentStringAsUtf16Chars.Add(utf16Char);

                            if (utf16Char == 0x00 /*'\0'*/)
                            {
                                // convert the utf16 char vector to a string
                                string utf16CharsAsString;
                                try
                                {
                                    utf16CharsAsString = new string(currentStringAsUtf16Chars.ToArray(), 0, currentStringAsUtf16Chars.Count - 1);
                                }
                                catch
                                {
                                    return Result.ErrorResult<IGetDevicePropertyValueError>(new IGetDevicePropertyValueError.StringDecodingError());
                                }

                                listOfStrings.Add(new IPnpDevicePropertyValue.String(utf16CharsAsString));

                                // reset our current string
                                currentStringAsUtf16Chars = new();
                            }
                        }
                        return Result.OkResult<IPnpDevicePropertyValue>(new IPnpDevicePropertyValue.ListOfValues(listOfStrings));
                    }
                    else
                    {
                        // single null-terminated string

                        if (propertyValueAsUtf16Chars[propertyValueAsUtf16Chars.Count - 1] == '\0')
                        {
                            // this is the correct terminator value; proceed
                        }
                        else
                        {
                            return Result.ErrorResult<IGetDevicePropertyValueError>(new IGetDevicePropertyValueError.StringTerminationError());
                        }

                        // convert the utf16 char vector to a string
                        string propertyBufferAsString;
                        try
                        {
                            propertyBufferAsString = new string(propertyValueAsUtf16Chars.ToArray(), 0, propertyValueAsUtf16Chars.Count - 1);
                        }
                        catch
                        {
                            return Result.ErrorResult<IGetDevicePropertyValueError>(new IGetDevicePropertyValueError.StringDecodingError());
                        }

                        return Result.OkResult<IPnpDevicePropertyValue>(new IPnpDevicePropertyValue.String(propertyBufferAsString));
                    }
                }
            case (uint)Windows.Win32.Devices.Properties.DEVPROPTYPE.DEVPROP_TYPE_UINT16:
                {
                    const int fixedSizeOfValueType = 2;

                    Func<byte[], IPnpDevicePropertyValue> bufferToPnpDevicePropertyValueFunc = delegate (byte[] buffer)
                    {
                        // convert the byte array to a uint16 (using native endian)
                        ushort value = BitConverter.ToUInt16(buffer, 0);
                        return new IPnpDevicePropertyValue.UInt16(value);
                    };

                    return createReturnValueForFixedSizePropertyFunc(fixedSizeOfValueType, propertyValueIsArray, bufferToPnpDevicePropertyValueFunc);
                }
            case (uint)Windows.Win32.Devices.Properties.DEVPROPTYPE.DEVPROP_TYPE_UINT32:
                {
                    const int fixedSizeOfValueType = 4;

                    Func<byte[], IPnpDevicePropertyValue> bufferToPnpDevicePropertyValueFunc = delegate (byte[] buffer)
                    {
                        // convert the byte array to a uint32 (using native endian)
                        uint value = BitConverter.ToUInt32(buffer, 0);
                        return new IPnpDevicePropertyValue.UInt32(value);
                    };

                    return createReturnValueForFixedSizePropertyFunc(fixedSizeOfValueType, propertyValueIsArray, bufferToPnpDevicePropertyValueFunc);
                }
            default:
                return Result.OkResult<IPnpDevicePropertyValue>(new IPnpDevicePropertyValue.UnsupportedPropertyDataType((uint)propertyType));
        }
    }

    // NOTE: this function calculates a mask which will fit any value equal to or less than the supplied value; if the value is not a power of two (minus one)...then the mask will also cover any numbers up to the next power of two (minus one)
    private static uint CalculateMaskToFitValue(uint value)
    {
        var numberOfHighZeroBits = 0;
        for (var index = 0; index <= 31; index += 1)
        {
            if ((value & (1 << (31 - index))) != 0)
            {
                break;
            }

            numberOfHighZeroBits += 1;
        }

        var result = 0xFFFF_FFFFu >> numberOfHighZeroBits;

        return result;
    }
}
