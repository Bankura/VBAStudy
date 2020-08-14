Attribute VB_Name = "WMIEx"
Option Explicit

'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/
'*
'* [機能名] WbemScriptingラップ・拡張モジュール
'* [詳  細] WbemScriptingのWrapperとしての機能を提供する他、
'*          WbemScriptingを使用したユーティリティを提供する。
'*          ラップするWbemScriptingライブラリは以下のものとする。
'*              [name] Microsoft WMI Scripting V1.2 Library
'*              [dll] C:\Windows\System32\wbem\wbemdisp.TLB
'* [参  考]
'*  <https://msdn.microsoft.com/ja-jp/windows/aa393259(v=vs.80)>
'*
'* @author Bankura
'* Copyright (c) 2019 Bankura
'*/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/_/

'******************************************************************************
'* Enum定義
'******************************************************************************

'*-----------------------------------------------------------------------------
'* WbemAuthenticationLevelEnum
'* Defines the security authentication level.
'*-----------------------------------------------------------------------------
Public Enum WbemAuthenticationLevelEnum
    wbemAuthenticationLevelCall = 3
    wbemAuthenticationLevelConnect = 2
    wbemAuthenticationLevelDefault = 0
    wbemAuthenticationLevelNone = 1
    wbemAuthenticationLevelPkt = 4
    wbemAuthenticationLevelPktIntegrity = 5
    wbemAuthenticationLevelPktPrivacy = 6
End Enum

'*-----------------------------------------------------------------------------
'* WbemChangeFlagEnum
'* Defines semantics of putting a Class or Instance.
'*-----------------------------------------------------------------------------
Public Enum WbemChangeFlagEnum
    wbemChangeFlagAdvisory = 65536
    wbemChangeFlagCreateOnly = 2
    wbemChangeFlagCreateOrUpdate = 0
    wbemChangeFlagStrongValidation = 128
    wbemChangeFlagUpdateCompatible = 0
    wbemChangeFlagUpdateForceMode = 64
    wbemChangeFlagUpdateOnly = 1
    wbemChangeFlagUpdateSafeMode = 32
End Enum

'*-----------------------------------------------------------------------------
'* WbemCimtypeEnum
'* Defines the valid CIM Types of a Property value.
'*-----------------------------------------------------------------------------
Public Enum WbemCimtypeEnum
    wbemCimtypeBoolean = 11
    wbemCimtypeChar16 = 103
    wbemCimtypeDatetime = 101
    wbemCimtypeObject = 13
    wbemCimtypeReal32 = 4
    wbemCimtypeReal64 = 5
    wbemCimtypeReference = 102
    wbemCimtypeSint16 = 2
    wbemCimtypeSint32 = 3
    wbemCimtypeSint64 = 20
    wbemCimtypeSint8 = 16
    wbemCimtypeString = 8
    wbemCimtypeUint16 = 18
    wbemCimtypeUint32 = 19
    wbemCimtypeUint64 = 21
    wbemCimtypeUint8 = 17
End Enum

'*-----------------------------------------------------------------------------
'* WbemComparisonFlagEnum
'* Defines settings for object comparison.
'*-----------------------------------------------------------------------------
Public Enum WbemComparisonFlagEnum
    wbemComparisonFlagIgnoreCase = 16
    wbemComparisonFlagIgnoreClass = 8
    wbemComparisonFlagIgnoreDefaultValues = 4
    wbemComparisonFlagIgnoreFlavor = 32
    wbemComparisonFlagIgnoreObjectSource = 2
    wbemComparisonFlagIgnoreQualifiers = 1
    wbemComparisonFlagIncludeAll = 0
End Enum

'*-----------------------------------------------------------------------------
'* WbemConnectOptionsEnum
'* Used to define connection behavior.
'*-----------------------------------------------------------------------------
Public Enum WbemConnectOptionsEnum
    wbemConnectFlagUseMaxWait = 128
End Enum

'*-----------------------------------------------------------------------------
'* WbemErrorEnum
'* Defines the errors that may be returned by the WBEM Scripting library.
'*-----------------------------------------------------------------------------
Public Enum WbemErrorEnum
    wbemErrAccessDenied = -2147217405
    wbemErrAggregatingByObject = -2147217315
    wbemErrAlreadyExists = -2147217383
    wbemErrAmbiguousOperation = -2147217301
    wbemErrAmendedObject = -2147217306
    wbemErrBackupRestoreWinmgmtRunning = -2147217312
    wbemErrBufferTooSmall = -2147217348
    wbemErrCallCancelled = -2147217358
    wbemErrCannotBeAbstract = -2147217307
    wbemErrCannotBeKey = -2147217377
    wbemErrCannotBeSingleton = -2147217364
    wbemErrCannotChangeIndexInheritance = -2147217328
    wbemErrCannotChangeKeyInheritance = -2147217335
    wbemErrCircularReference = -2147217337
    wbemErrClassHasChildren = -2147217371
    wbemErrClassHasInstances = -2147217370
    wbemErrClassNameTooWide = -2147217292
    wbemErrClientTooSlow = -2147217305
    wbemErrConnectionFailed = -2147217295
    wbemErrCriticalError = -2147217398
    wbemErrDatabaseVerMismatch = -2147217288
    wbemErrEncryptedConnectionRequired = -2147217273
    wbemErrFailed = -2147217407
    wbemErrFatalTransportError = -2147217274
    wbemErrForcedRollback = -2147217298
    wbemErrHandleOutOfDate = -2147217296
    wbemErrIllegalNull = -2147217368
    wbemErrIllegalOperation = -2147217378
    wbemErrIncompleteClass = -2147217376
    wbemErrInitializationFailure = -2147217388
    wbemErrInvalidAssociation = -2147217302
    wbemErrInvalidCimType = -2147217363
    wbemErrInvalidClass = -2147217392
    wbemErrInvalidContext = -2147217401
    wbemErrInvalidDuplicateParameter = -2147217341
    wbemErrInvalidFlavor = -2147217338
    wbemErrInvalidHandleRequest = -2147217294
    wbemErrInvalidLocale = -2147217280
    wbemErrInvalidMethod = -2147217362
    wbemErrInvalidMethodParameters = -2147217361
    wbemErrInvalidNamespace = -2147217394
    wbemErrInvalidObject = -2147217393
    wbemErrInvalidObjectPath = -2147217350
    wbemErrInvalidOperation = -2147217386
    wbemErrInvalidOperator = -2147217309
    wbemErrInvalidParameter = -2147217400
    wbemErrInvalidParameterId = -2147217353
    wbemErrInvalidProperty = -2147217359
    wbemErrInvalidPropertyType = -2147217366
    wbemErrInvalidProviderRegistration = -2147217390
    wbemErrInvalidQualifier = -2147217342
    wbemErrInvalidQualifierType = -2147217367
    wbemErrInvalidQuery = -2147217385
    wbemErrInvalidQueryType = -2147217384
    wbemErrInvalidStream = -2147217397
    wbemErrInvalidSuperclass = -2147217395
    wbemErrInvalidSyntax = -2147217375
    wbemErrLocalCredentials = -2147217308
    wbemErrMarshalInvalidSignature = -2147217343
    wbemErrMarshalVersionMismatch = -2147217344
    wbemErrMethodDisabled = -2147217322
    wbemErrMethodNameTooWide = -2147217291
    wbemErrMethodNotImplemented = -2147217323
    wbemErrMissingAggregationList = -2147217317
    wbemErrMissingGroupWithin = -2147217318
    wbemErrMissingParameter = -2147217354
    wbemErrNonConsecutiveParameterIds = -2147217352
    wbemErrNondecoratedObject = -2147217374
    wbemErrNoSchema = -2147217277
    wbemErrNotAvailable = -2147217399
    wbemErrNotEventClass = -2147217319
    wbemErrNotFound = -2147217406
    wbemErrNotSupported = -2147217396
    wbemErrNullSecurityDescriptor = -2147217304
    wbemErrOutOfDiskSpace = -2147217349
    wbemErrOutOfMemory = -2147217402
    wbemErrOverrideNotAllowed = -2147217382
    wbemErrParameterIdOnRetval = -2147217351
    wbemErrPrivilegeNotHeld = -2147217310
    wbemErrPropagatedMethod = -2147217356
    wbemErrPropagatedProperty = -2147217380
    wbemErrPropagatedQualifier = -2147217381
    wbemErrPropertyNameTooWide = -2147217293
    wbemErrPropertyNotAnObject = -2147217316
    wbemErrProviderAlreadyRegistered = -2147217276
    wbemErrProviderFailure = -2147217404
    wbemErrProviderLoadFailure = -2147217389
    wbemErrProviderNotCapable = -2147217372
    wbemErrProviderNotFound = -2147217391
    wbemErrProviderNotRegistered = -2147217275
    wbemErrProviderSuspended = -2147217279
    wbemErrQualifierNameTooWide = -2147217290
    wbemErrQueryNotImplemented = -2147217369
    wbemErrQueueOverflow = -2147217311
    wbemErrQuotaViolation = -2147217300
    wbemErrReadOnly = -2147217373
    wbemErrRefresherBusy = -2147217321
    wbemErrRegistrationTooBroad = -2147213311
    wbemErrRegistrationTooPrecise = -2147213310
    wbemErrRerunCommand = -2147217289
    wbemErrResetToDefault = -2147209214
    wbemErrServerTooBusy = -2147217339
    wbemErrShuttingDown = -2147217357
    wbemErrSynchronizationRequired = -2147217278
    wbemErrSystemProperty = -2147217360
    wbemErrTimedout = -2147209215
    wbemErrTimeout = -2147217303
    wbemErrTooManyProperties = -2147217327
    wbemErrTooMuchData = -2147217340
    wbemErrTransactionConflict = -2147217299
    wbemErrTransportFailure = -2147217387
    wbemErrTypeMismatch = -2147217403
    wbemErrUnexpected = -2147217379
    wbemErrUninterpretableProviderQuery = -2147217313
    wbemErrUnknownObjectType = -2147217346
    wbemErrUnknownPacketType = -2147217345
    wbemErrUnparsableQuery = -2147217320
    wbemErrUnsupportedClassUpdate = -2147217336
    wbemErrUnsupportedLocale = -2147217297
    wbemErrUnsupportedParameter = -2147217355
    wbemErrUnsupportedPutExtension = -2147217347
    wbemErrUpdateOverrideNotAllowed = -2147217325
    wbemErrUpdatePropagatedMethod = -2147217324
    wbemErrUpdateTypeMismatch = -2147217326
    wbemErrValueOutOfRange = -2147217365
    wbemErrVetoPut = -2147217287
    wbemNoErr = 0
End Enum

'*-----------------------------------------------------------------------------
'* WbemFlagEnum
'* Defines behavior of various interface calls.
'*-----------------------------------------------------------------------------
Public Enum WbemFlagEnum
    wbemFlagBidirectional = 0
    wbemFlagDirectRead = 512
    wbemFlagDontSendStatus = 0
    wbemFlagEnsureLocatable = 256
    wbemFlagForwardOnly = 32
    wbemFlagGetDefault = 0
    wbemFlagNoErrorObject = 64
    wbemFlagReturnErrorObject = 0
    wbemFlagReturnImmediately = 16
    wbemFlagReturnWhenComplete = 0
    wbemFlagSendOnlySelected = 0
    wbemFlagSendStatus = 128
    wbemFlagSpawnInstance = 1
    wbemFlagUseAmendedQualifiers = 131072
    wbemFlagUseCurrentTime = 1
End Enum

'*-----------------------------------------------------------------------------
'* WbemImpersonationLevelEnum
'* Defines the security impersonation level.
'*-----------------------------------------------------------------------------
Public Enum WbemImpersonationLevelEnum
    wbemImpersonationLevelAnonymous = 1
    wbemImpersonationLevelDelegate = 4
    wbemImpersonationLevelIdentify = 2
    wbemImpersonationLevelImpersonate = 3
End Enum

'*-----------------------------------------------------------------------------
'* WbemObjectTextFormatEnum
'* Defines object text formats.
'*-----------------------------------------------------------------------------
Public Enum WbemObjectTextFormatEnum
    wbemObjectTextFormatCIMDTD20 = 1
    wbemObjectTextFormatWMIDTD20 = 2
End Enum

'*-----------------------------------------------------------------------------
'* WbemPrivilegeEnum
'* Defines a privilege.
'*-----------------------------------------------------------------------------
Public Enum WbemPrivilegeEnum
    wbemPrivilegeAudit = 20
    wbemPrivilegeBackup = 16
    wbemPrivilegeChangeNotify = 22
    wbemPrivilegeCreatePagefile = 14
    wbemPrivilegeCreatePermanent = 15
    wbemPrivilegeCreateToken = 1
    wbemPrivilegeDebug = 19
    wbemPrivilegeEnableDelegation = 26
    wbemPrivilegeIncreaseBasePriority = 13
    wbemPrivilegeIncreaseQuota = 4
    wbemPrivilegeLoadDriver = 9
    wbemPrivilegeLockMemory = 3
    wbemPrivilegeMachineAccount = 5
    wbemPrivilegeManageVolume = 27
    wbemPrivilegePrimaryToken = 2
    wbemPrivilegeProfileSingleProcess = 12
    wbemPrivilegeRemoteShutdown = 23
    wbemPrivilegeRestore = 17
    wbemPrivilegeSecurity = 7
    wbemPrivilegeShutdown = 18
    wbemPrivilegeSyncAgent = 25
    wbemPrivilegeSystemEnvironment = 21
    wbemPrivilegeSystemProfile = 10
    wbemPrivilegeSystemtime = 11
    wbemPrivilegeTakeOwnership = 8
    wbemPrivilegeTcb = 6
    wbemPrivilegeUndock = 24
End Enum

'*-----------------------------------------------------------------------------
'* WbemQueryFlagEnum
'* Defines depth of enumeration or query.
'*-----------------------------------------------------------------------------
Public Enum WbemQueryFlagEnum
    wbemQueryFlagDeep = 0
    wbemQueryFlagPrototype = 2
    wbemQueryFlagShallow = 1
End Enum

'*-----------------------------------------------------------------------------
'* WbemTextFlagEnum
'* Defines content of generated object text.
'*-----------------------------------------------------------------------------
Public Enum WbemTextFlagEnum
    wbemTextFlagNoFlavors = 1
End Enum

'*-----------------------------------------------------------------------------
'* WbemTimeout
'* Defines timeout constants.
'*-----------------------------------------------------------------------------
Public Enum WbemTimeout
    wbemTimeoutInfinite = -1
End Enum

'*-----------------------------------------------------------------------------
'* 拡張Enum定義
'*-----------------------------------------------------------------------------

'*-----------------------------------------------------------------------------
'* WmiClassEnum
'* Wmiで扱うClassのEnum定義です。
'*-----------------------------------------------------------------------------
Public Enum WmiClassEnum
    wmiEnumClassCIMAction
    wmiEnumClassCIMActionSequence
    wmiEnumClassCIMActsAsSpare
    wmiEnumClassCIMAdjacentSlots
    wmiEnumClassCIMAggregatePExtent
    wmiEnumClassCIMAggregatePSExtent
    wmiEnumClassCIMAggregateRedundancyComponent
    wmiEnumClassCIMAlarmDevice
    wmiEnumClassCIMAllocatedResource
    wmiEnumClassCIMApplicationSystem
    wmiEnumClassCIMApplicationSystemSoftwareFeature
    wmiEnumClassCIMAssociatedAlarm
    wmiEnumClassCIMAssociatedBattery
    wmiEnumClassCIMAssociatedCooling
    wmiEnumClassCIMAssociatedMemory
    wmiEnumClassCIMAssociatedProcessorMemory
    wmiEnumClassCIMAssociatedSensor
    wmiEnumClassCIMAssociatedSupplyCurrentSensor
    wmiEnumClassCIMAssociatedSupplyVoltageSensor
    wmiEnumClassCIMBasedOn
    wmiEnumClassCIMBattery
    wmiEnumClassCIMBinarySensor
    wmiEnumClassCIMBIOSElement
    wmiEnumClassCIMBIOSFeature
    wmiEnumClassCIMBIOSFeatureBIOSElements
    wmiEnumClassCIMBIOSLoadedInNV
    wmiEnumClassCIMBootOSFromFS
    wmiEnumClassCIMBootSAP
    wmiEnumClassCIMBootService
    wmiEnumClassCIMBootServiceAccessBySAP
    wmiEnumClassCIMCacheMemory
    wmiEnumClassCIMCard
    wmiEnumClassCIMCardInSlot
    wmiEnumClassCIMCardOnCard
    wmiEnumClassCIMCDROMDrive
    wmiEnumClassCIMChassis
    wmiEnumClassCIMChassisInRack
    wmiEnumClassCIMCheck
    wmiEnumClassCIMChip
    wmiEnumClassCIMClusteringSAP
    wmiEnumClassCIMClusteringService
    wmiEnumClassCIMClusterServiceAccessBySAP
    wmiEnumClassCIMCollectedCollections
    wmiEnumClassCIMCollectedMSEs
    wmiEnumClassCIMCollectionOfMSEs
    wmiEnumClassCIMCollectionOfSensors
    wmiEnumClassCIMCollectionSetting
    wmiEnumClassCIMCompatibleProduct
    wmiEnumClassCIMComponent
    wmiEnumClassCIMComputerSystem
    wmiEnumClassCIMComputerSystemDMA
    wmiEnumClassCIMComputerSystemIRQ
    wmiEnumClassCIMComputerSystemMappedIO
    wmiEnumClassCIMComputerSystemPackage
    wmiEnumClassCIMComputerSystemResource
    wmiEnumClassCIMConfiguration
    wmiEnumClassCIMConnectedTo
    wmiEnumClassCIMConnectorOnPackage
    wmiEnumClassCIMContainer
    wmiEnumClassCIMControlledBy
    wmiEnumClassCIMController
    wmiEnumClassCIMCoolingDevice
    wmiEnumClassCIMCopyFileAction
    wmiEnumClassCIMCreateDirectoryAction
    wmiEnumClassCIMCurrentSensor
    wmiEnumClassCIMDataFile
    wmiEnumClassCIMDependency
    wmiEnumClassCIMDependencyContext
    wmiEnumClassCIMDesktopMonitor
    wmiEnumClassCIMDeviceAccessedByFile
    wmiEnumClassCIMDeviceConnection
    wmiEnumClassCIMDeviceErrorCounts
    wmiEnumClassCIMDeviceFile
    wmiEnumClassCIMDeviceSAPImplementation
    wmiEnumClassCIMDeviceServiceImplementation
    wmiEnumClassCIMDeviceSoftware
    wmiEnumClassCIMDirectory
    wmiEnumClassCIMDirectoryAction
    wmiEnumClassCIMDirectoryContainsFile
    wmiEnumClassCIMDirectorySpecification
    wmiEnumClassCIMDirectorySpecificationFile
    wmiEnumClassCIMDiscreteSensor
    wmiEnumClassCIMDiskDrive
    wmiEnumClassCIMDisketteDrive
    wmiEnumClassCIMDiskPartition
    wmiEnumClassCIMDiskSpaceCheck
    wmiEnumClassCIMDisplay
    wmiEnumClassCIMDMA
    wmiEnumClassCIMDocked
    wmiEnumClassCIMElementCapacity
    wmiEnumClassCIMElementConfiguration
    wmiEnumClassCIMElementSetting
    wmiEnumClassCIMElementsLinked
    wmiEnumClassCIMErrorCountersForDevice
    wmiEnumClassCIMExecuteProgram
    wmiEnumClassCIMExport
    wmiEnumClassCIMExtraCapacityGroup
    wmiEnumClassCIMFan
    wmiEnumClassCIMFileAction
    wmiEnumClassCIMFileSpecification
    wmiEnumClassCIMFileStorage
    wmiEnumClassCIMFileSystem
    wmiEnumClassCIMFlatPanel
    wmiEnumClassCIMFromDirectoryAction
    wmiEnumClassCIMFromDirectorySpecification
    wmiEnumClassCIMFRU
    wmiEnumClassCIMFRUIncludesProduct
    wmiEnumClassCIMFRUPhysicalElements
    wmiEnumClassCIMHeatPipe
    wmiEnumClassCIMHostedAccessPoint
    wmiEnumClassCIMHostedBootSAP
    wmiEnumClassCIMHostedBootService
    wmiEnumClassCIMHostedFileSystem
    wmiEnumClassCIMHostedJobDestination
    wmiEnumClassCIMHostedService
    wmiEnumClassCIMInfraredController
    wmiEnumClassCIMInstalledOS
    wmiEnumClassCIMInstalledSoftwareElement
    wmiEnumClassCIMIRQ
    wmiEnumClassCIMJob
    wmiEnumClassCIMJobDestination
    wmiEnumClassCIMJobDestinationJobs
    wmiEnumClassCIMKeyboard
    wmiEnumClassCIMLinkHasConnector
    wmiEnumClassCIMLocalFileSystem
    wmiEnumClassCIMLocation
    wmiEnumClassCIMLogicalDevice
    wmiEnumClassCIMLogicalDisk
    wmiEnumClassCIMLogicalDiskBasedOnPartition
    wmiEnumClassCIMLogicalDiskBasedOnVolumeSet
    wmiEnumClassCIMLogicalElement
    wmiEnumClassCIMLogicalFile
    wmiEnumClassCIMLogicalIdentity
    wmiEnumClassCIMMagnetoOpticalDrive
    wmiEnumClassCIMManagedSystemElement
    wmiEnumClassCIMManagementController
    wmiEnumClassCIMMediaAccessDevice
    wmiEnumClassCIMMediaPresent
    wmiEnumClassCIMMemory
    wmiEnumClassCIMMemoryCapacity
    wmiEnumClassCIMMemoryCheck
    wmiEnumClassCIMMemoryMappedIO
    wmiEnumClassCIMMemoryOnCard
    wmiEnumClassCIMMemoryWithMedia
    wmiEnumClassCIMModifySettingAction
    wmiEnumClassCIMMonitorResolution
    wmiEnumClassCIMMonitorSetting
    wmiEnumClassCIMMount
    wmiEnumClassCIMMultiStateSensor
    wmiEnumClassCIMNetworkAdapter
    wmiEnumClassCIMNFS
    wmiEnumClassCIMNonVolatileStorage
    wmiEnumClassCIMNumericSensor
    wmiEnumClassCIMOperatingSystem
    wmiEnumClassCIMOperatingSystemSoftwareFeature
    wmiEnumClassCIMOSProcess
    wmiEnumClassCIMOSVersionCheck
    wmiEnumClassCIMPackageAlarm
    wmiEnumClassCIMPackageCooling
    wmiEnumClassCIMPackagedComponent
    wmiEnumClassCIMPackageInChassis
    wmiEnumClassCIMPackageInSlot
    wmiEnumClassCIMPackageTempSensor
    wmiEnumClassCIMParallelController
    wmiEnumClassCIMParticipatesInSet
    wmiEnumClassCIMPCIController
    wmiEnumClassCIMPCMCIAController
    wmiEnumClassCIMPCVideoController
    wmiEnumClassCIMPExtentRedundancyComponent
    wmiEnumClassCIMPhysicalCapacity
    wmiEnumClassCIMPhysicalComponent
    wmiEnumClassCIMPhysicalConnector
    wmiEnumClassCIMPhysicalElement
    wmiEnumClassCIMPhysicalElementLocation
    wmiEnumClassCIMPhysicalExtent
    wmiEnumClassCIMPhysicalFrame
    wmiEnumClassCIMPhysicalLink
    wmiEnumClassCIMPhysicalMedia
    wmiEnumClassCIMPhysicalMemory
    wmiEnumClassCIMPhysicalPackage
    wmiEnumClassCIMPointingDevice
    wmiEnumClassCIMPotsModem
    wmiEnumClassCIMPowerSupply
    wmiEnumClassCIMPrinter
    wmiEnumClassCIMProcess
    wmiEnumClassCIMProcessExecutable
    wmiEnumClassCIMProcessor
    wmiEnumClassCIMProcessThread
    wmiEnumClassCIMProduct
    wmiEnumClassCIMProductFRU
    wmiEnumClassCIMProductParentChild
    wmiEnumClassCIMProductPhysicalElements
    wmiEnumClassCIMProductProductDependency
    wmiEnumClassCIMProductSoftwareFeatures
    wmiEnumClassCIMProductSupport
    wmiEnumClassCIMProtectedSpaceExtent
    wmiEnumClassCIMPSExtentBasedOnPExtent
    wmiEnumClassCIMRack
    wmiEnumClassCIMRealizes
    wmiEnumClassCIMRealizesAggregatePExtent
    wmiEnumClassCIMRealizesDiskPartition
    wmiEnumClassCIMRealizesPExtent
    wmiEnumClassCIMRebootAction
    wmiEnumClassCIMRedundancyComponent
    wmiEnumClassCIMRedundancyGroup
    wmiEnumClassCIMRefrigeration
    wmiEnumClassCIMRelatedStatistics
    wmiEnumClassCIMRemoteFileSystem
    wmiEnumClassCIMRemoveDirectoryAction
    wmiEnumClassCIMRemoveFileAction
    wmiEnumClassCIMReplacementSet
    wmiEnumClassCIMResidesOnExtent
    wmiEnumClassCIMRunningOS
    wmiEnumClassCIMSAPSAPDependency
    wmiEnumClassCIMScanner
    wmiEnumClassCIMSCSIController
    wmiEnumClassCIMSCSIInterface
    wmiEnumClassCIMSensor
    wmiEnumClassCIMSerialController
    wmiEnumClassCIMSerialInterface
    wmiEnumClassCIMService
    wmiEnumClassCIMServiceAccessBySAP
    wmiEnumClassCIMServiceAccessPoint
    wmiEnumClassCIMServiceSAPDependency
    wmiEnumClassCIMServiceServiceDependency
    wmiEnumClassCIMSetting
    wmiEnumClassCIMSettingCheck
    wmiEnumClassCIMSettingContext
    wmiEnumClassCIMSlot
    wmiEnumClassCIMSlotInSlot
    wmiEnumClassCIMSoftwareElement
    wmiEnumClassCIMSoftwareElementActions
    wmiEnumClassCIMSoftwareElementChecks
    wmiEnumClassCIMSoftwareElementVersionCheck
    wmiEnumClassCIMSoftwareFeature
    wmiEnumClassCIMSoftwareFeatureSAPImplementation
    wmiEnumClassCIMSoftwareFeatureServiceImplementation
    wmiEnumClassCIMSoftwareFeatureSoftwareElements
    wmiEnumClassCIMSpareGroup
    wmiEnumClassCIMStatisticalInformation
    wmiEnumClassCIMStatistics
    wmiEnumClassCIMStorageDefect
    wmiEnumClassCIMStorageError
    wmiEnumClassCIMStorageExtent
    wmiEnumClassCIMStorageRedundancyGroup
    wmiEnumClassCIMSupportAccess
    wmiEnumClassCIMSwapSpaceCheck
    wmiEnumClassCIMSystem
    wmiEnumClassCIMSystemComponent
    wmiEnumClassCIMSystemDevice
    wmiEnumClassCIMSystemResource
    wmiEnumClassCIMTachometer
    wmiEnumClassCIMTapeDrive
    wmiEnumClassCIMTemperatureSensor
    wmiEnumClassCIMThread
    wmiEnumClassCIMToDirectoryAction
    wmiEnumClassCIMToDirectorySpecification
    wmiEnumClassCIMUninterruptiblePowerSupply
    wmiEnumClassCIMUnitaryComputerSystem
    wmiEnumClassCIMUSBController
    wmiEnumClassCIMUSBControllerHasHub
    wmiEnumClassCIMUserDevice
    wmiEnumClassCIMVersionCompatibilityCheck
    wmiEnumClassCIMVideoBIOSElement
    wmiEnumClassCIMVideoBIOSFeature
    wmiEnumClassCIMVideoBIOSFeatureVideoBIOSElements
    wmiEnumClassCIMVideoController
    wmiEnumClassCIMVideoControllerResolution
    wmiEnumClassCIMVideoSetting
    wmiEnumClassCIMVolatileStorage
    wmiEnumClassCIMVoltageSensor
    wmiEnumClassCIMVolumeSet
    wmiEnumClassCIMWORMDrive
    wmiEnumClassMSFTNCProvAccessCheck
    wmiEnumClassMSFTNCProvCancelQuery
    wmiEnumClassMSFTNCProvClientConnected
    wmiEnumClassMSFTNCProvEvent
    wmiEnumClassMSFTNCProvNewQuery
    wmiEnumClassMSFTNetBadAccount
    wmiEnumClassMSFTNetBadServiceState
    wmiEnumClassMSFTNetBootSystemDriversFailed
    wmiEnumClassMSFTNetCallToFunctionFailed
    wmiEnumClassMSFTNetCallToFunctionFailedII
    wmiEnumClassMSFTNetCircularDependencyAuto
    wmiEnumClassMSFTNetCircularDependencyDemand
    wmiEnumClassMSFTNetConnectionTimeout
    wmiEnumClassMSFTNetDependOnLaterGroup
    wmiEnumClassMSFTNetDependOnLaterService
    wmiEnumClassMSFTNetFirstLogonFailed
    wmiEnumClassMSFTNetFirstLogonFailedII
    wmiEnumClassMSFTNetReadfileTimeout
    wmiEnumClassMSFTNetRevertedToLastKnownGood
    wmiEnumClassMSFTNetServiceConfigBackoutFailed
    wmiEnumClassMSFTNetServiceControlSuccess
    wmiEnumClassMSFTNetServiceCrash
    wmiEnumClassMSFTNetServiceCrashNoAction
    wmiEnumClassMSFTNetServiceDifferentPIDConnected
    wmiEnumClassMSFTNetServiceExitFailed
    wmiEnumClassMSFTNetServiceExitFailedSpecific
    wmiEnumClassMSFTNetServiceLogonTypeNotGranted
    wmiEnumClassMSFTNetServiceNotInteractive
    wmiEnumClassMSFTNetServiceRecoveryFailed
    wmiEnumClassMSFTNetServiceShutdownFailed
    wmiEnumClassMSFTNetServiceSlowStartup
    wmiEnumClassMSFTNetServiceStartFailed
    wmiEnumClassMSFTNetServiceStartFailedGroup
    wmiEnumClassMSFTNetServiceStartFailedII
    wmiEnumClassMSFTNetServiceStartFailedNone
    wmiEnumClassMSFTNetServiceStartHung
    wmiEnumClassMSFTNetServiceStartTypeChanged
    wmiEnumClassMSFTNetServiceStatusSuccess
    wmiEnumClassMSFTNetServiceStopControlSuccess
    wmiEnumClassMSFTNetSevereServiceFailed
    wmiEnumClassMSFTNetTakeOwnership
    wmiEnumClassMSFTNetTransactInvalid
    wmiEnumClassMSFTNetTransactTimeout
    wmiEnumClassMsftProviders
    wmiEnumClassMSFTSCMEvent
    wmiEnumClassMSFTSCMEventLogEvent
    wmiEnumClassMSFTWMIGenericNonCOMEvent
    wmiEnumClassMSFTWmiCancelNotificationSink
    wmiEnumClassMSFTWmiConsumerProviderEvent
    wmiEnumClassMSFTWmiConsumerProviderLoaded
    wmiEnumClassMSFTWmiConsumerProviderSinkLoaded
    wmiEnumClassMSFTWmiConsumerProviderSinkUnloaded
    wmiEnumClassMSFTWmiConsumerProviderUnloaded
    wmiEnumClassMSFTWmiEssEvent
    wmiEnumClassMSFTWmiFilterActivated
    wmiEnumClassMSFTWmiFilterDeactivated
    wmiEnumClassMSFTWmiFilterEvent
    wmiEnumClassMsftWmiProviderAccessCheckPost
    wmiEnumClassMsftWmiProviderAccessCheckPre
    wmiEnumClassMsftWmiProviderCancelQueryPost
    wmiEnumClassMsftWmiProviderCancelQueryPre
    wmiEnumClassMsftWmiProviderComServerLoadOperationEvent
    wmiEnumClassMsftWmiProviderComServerLoadOperationFailureEvent
    wmiEnumClassMsftWmiProviderCounters
    wmiEnumClassMsftWmiProviderCreateClassEnumAsyncEventPost
    wmiEnumClassMsftWmiProviderCreateClassEnumAsyncEventPre
    wmiEnumClassMsftWmiProviderCreateInstanceEnumAsyncEventPost
    wmiEnumClassMsftWmiProviderCreateInstanceEnumAsyncEventPre
    wmiEnumClassMsftWmiProviderDeleteClassAsyncEventPost
    wmiEnumClassMsftWmiProviderDeleteClassAsyncEventPre
    wmiEnumClassMsftWmiProviderDeleteInstanceAsyncEventPost
    wmiEnumClassMsftWmiProviderDeleteInstanceAsyncEventPre
    wmiEnumClassMsftWmiProviderExecMethodAsyncEventPost
    wmiEnumClassMsftWmiProviderExecMethodAsyncEventPre
    wmiEnumClassMsftWmiProviderExecQueryAsyncEventPost
    wmiEnumClassMsftWmiProviderExecQueryAsyncEventPre
    wmiEnumClassMsftWmiProviderGetObjectAsyncEventPost
    wmiEnumClassMsftWmiProviderGetObjectAsyncEventPre
    wmiEnumClassMsftWmiProviderInitializationOperationEvent
    wmiEnumClassMsftWmiProviderInitializationOperationFailureEvent
    wmiEnumClassMsftWmiProviderLoadOperationEvent
    wmiEnumClassMsftWmiProviderLoadOperationFailureEvent
    wmiEnumClassMsftWmiProviderNewQueryPost
    wmiEnumClassMsftWmiProviderNewQueryPre
    wmiEnumClassMsftWmiProviderOperationEvent
    wmiEnumClassMsftWmiProviderOperationEventPost
    wmiEnumClassMsftWmiProviderOperationEventPre
    wmiEnumClassMsftWmiProviderProvideEventsPost
    wmiEnumClassMsftWmiProviderProvideEventsPre
    wmiEnumClassMsftWmiProviderPutClassAsyncEventPost
    wmiEnumClassMsftWmiProviderPutClassAsyncEventPre
    wmiEnumClassMsftWmiProviderPutInstanceAsyncEventPost
    wmiEnumClassMsftWmiProviderPutInstanceAsyncEventPre
    wmiEnumClassMsftWmiProviderUnLoadOperationEvent
    wmiEnumClassMSFTWmiProviderEvent
    wmiEnumClassMSFTWmiRegisterNotificationSink
    wmiEnumClassMSFTWmiSelfEvent
    wmiEnumClassMSFTWmiThreadPoolEvent
    wmiEnumClassMSFTWmiThreadPoolThreadCreated
    wmiEnumClassMSFTWmiThreadPoolThreadDeleted
    wmiEnumClassWin321394Controller
    wmiEnumClassWin321394ControllerDevice
    wmiEnumClassWin32Account
    wmiEnumClassWin32AccountSID
    wmiEnumClassWin32ACE
    wmiEnumClassWin32ActionCheck
    wmiEnumClassWin32ActiveRoute
    wmiEnumClassWin32AllocatedResource
    wmiEnumClassWin32ApplicationCommandLine
    wmiEnumClassWin32ApplicationService
    wmiEnumClassWin32AssociatedProcessorMemory
    wmiEnumClassWin32AutochkSetting
    wmiEnumClassWin32BaseBoard
    wmiEnumClassWin32BaseService
    wmiEnumClassWin32Battery
    wmiEnumClassWin32Binary
    wmiEnumClassWin32BindImageAction
    wmiEnumClassWin32BIOS
    wmiEnumClassWin32BootConfiguration
    wmiEnumClassWin32Bus
    wmiEnumClassWin32CacheMemory
    wmiEnumClassWin32CDROMDrive
    wmiEnumClassWin32CheckCheck
    wmiEnumClassWin32CIMLogicalDeviceCIMDataFile
    wmiEnumClassWin32ClassicCOMApplicationClasses
    wmiEnumClassWin32ClassicCOMClass
    wmiEnumClassWin32ClassicCOMClassSetting
    wmiEnumClassWin32ClassicCOMClassSettings
    wmiEnumClassWin32ClassInfoAction
    wmiEnumClassWin32ClientApplicationSetting
    wmiEnumClassWin32ClusterShare
    wmiEnumClassWin32CodecFile
    wmiEnumClassWin32CollectionStatistics
    wmiEnumClassWin32COMApplication
    wmiEnumClassWin32COMApplicationClasses
    wmiEnumClassWin32COMApplicationSettings
    wmiEnumClassWin32COMClass
    wmiEnumClassWin32ComClassAutoEmulator
    wmiEnumClassWin32ComClassEmulator
    wmiEnumClassWin32CommandLineAccess
    wmiEnumClassWin32ComponentCategory
    wmiEnumClassWin32ComputerShutdownEvent
    wmiEnumClassWin32ComputerSystem
    wmiEnumClassWin32ComputerSystemEvent
    wmiEnumClassWin32ComputerSystemProcessor
    wmiEnumClassWin32ComputerSystemProduct
    wmiEnumClassWin32COMSetting
    wmiEnumClassWin32Condition
    wmiEnumClassWin32ConnectionShare
    wmiEnumClassWin32ControllerHasHub
    wmiEnumClassWin32CreateFolderAction
    wmiEnumClassWin32CurrentProbe
    wmiEnumClassWin32CurrentTime
    wmiEnumClassWin32DCOMApplication
    wmiEnumClassWin32DCOMApplicationAccessAllowedSetting
    wmiEnumClassWin32DCOMApplicationLaunchAllowedSetting
    wmiEnumClassWin32DCOMApplicationSetting
    wmiEnumClassWin32DependentService
    wmiEnumClassWin32Desktop
    wmiEnumClassWin32DesktopMonitor
    wmiEnumClassWin32DeviceBus
    wmiEnumClassWin32DeviceChangeEvent
    wmiEnumClassWin32DeviceMemoryAddress
    wmiEnumClassWin32DeviceSettings
    wmiEnumClassWin32DfsNode
    wmiEnumClassWin32DfsNodeTarget
    wmiEnumClassWin32DfsTarget
    wmiEnumClassWin32Directory
    wmiEnumClassWin32DirectorySpecification
    wmiEnumClassWin32DiskDrive
    wmiEnumClassWin32DiskDrivePhysicalMedia
    wmiEnumClassWin32DiskDriveToDiskPartition
    wmiEnumClassWin32DiskPartition
    wmiEnumClassWin32DiskQuota
    wmiEnumClassWin32DisplayConfiguration
    wmiEnumClassWin32DisplayControllerConfiguration
    wmiEnumClassWin32DMAChannel
    wmiEnumClassWin32DriverForDevice
    wmiEnumClassWin32DuplicateFileAction
    wmiEnumClassWin32Environment
    wmiEnumClassWin32EnvironmentSpecification
    wmiEnumClassWin32ExtensionInfoAction
    wmiEnumClassWin32Fan
    wmiEnumClassWin32FileSpecification
    wmiEnumClassWin32FloppyController
    wmiEnumClassWin32FloppyDrive
    wmiEnumClassWin32FolderRedirection
    wmiEnumClassWin32FolderRedirectionHealth
    wmiEnumClassWin32FolderRedirectionHealthConfiguration
    wmiEnumClassWin32FolderRedirectionUserConfiguration
    wmiEnumClassWin32FontInfoAction
    wmiEnumClassWin32Group
    wmiEnumClassWin32GroupInDomain
    wmiEnumClassWin32GroupUser
    wmiEnumClassWin32HeatPipe
    wmiEnumClassWin32IDEController
    wmiEnumClassWin32IDEControllerDevice
    wmiEnumClassWin32ImplementedCategory
    wmiEnumClassWin32InfraredDevice
    wmiEnumClassWin32IniFileSpecification
    wmiEnumClassWin32InstalledProgramFramework
    wmiEnumClassWin32InstalledSoftwareElement
    wmiEnumClassWin32InstalledStoreProgram
    wmiEnumClassWin32InstalledWin32Program
    wmiEnumClassWin32IP4PersistedRouteTable
    wmiEnumClassWin32IP4RouteTable
    wmiEnumClassWin32IP4RouteTableEvent
    wmiEnumClassWin32IRQResource
    wmiEnumClassWin32JobObjectStatus
    wmiEnumClassWin32Keyboard
    wmiEnumClassWin32LaunchCondition
    wmiEnumClassWin32LoadOrderGroup
    wmiEnumClassWin32LoadOrderGroupServiceDependencies
    wmiEnumClassWin32LoadOrderGroupServiceMembers
    wmiEnumClassWin32LocalTime
    wmiEnumClassWin32LoggedOnUser
    wmiEnumClassWin32LogicalDisk
    wmiEnumClassWin32LogicalDiskRootDirectory
    wmiEnumClassWin32LogicalDiskToPartition
    wmiEnumClassWin32LogicalFileAccess
    wmiEnumClassWin32LogicalFileAuditing
    wmiEnumClassWin32LogicalFileGroup
    wmiEnumClassWin32LogicalFileOwner
    wmiEnumClassWin32LogicalFileSecuritySetting
    wmiEnumClassWin32LogicalProgramGroup
    wmiEnumClassWin32LogicalProgramGroupDirectory
    wmiEnumClassWin32LogicalProgramGroupItem
    wmiEnumClassWin32LogicalProgramGroupItemDataFile
    wmiEnumClassWin32LogicalShareAccess
    wmiEnumClassWin32LogicalShareAuditing
    wmiEnumClassWin32LogicalShareSecuritySetting
    wmiEnumClassWin32LogonSession
    wmiEnumClassWin32LogonSessionMappedDisk
    wmiEnumClassWin32LUID
    wmiEnumClassWin32LUIDandAttributes
    wmiEnumClassWin32ManagedSystemElementResource
    wmiEnumClassWin32MappedLogicalDisk
    wmiEnumClassWin32MemoryArray
    wmiEnumClassWin32MemoryArrayLocation
    wmiEnumClassWin32MemoryDevice
    wmiEnumClassWin32MemoryDeviceArray
    wmiEnumClassWin32MemoryDeviceLocation
    wmiEnumClassWin32MethodParameterClass
    wmiEnumClassWin32MIMEInfoAction
    wmiEnumClassWin32ModuleLoadTrace
    wmiEnumClassWin32ModuleTrace
    wmiEnumClassWin32MotherboardDevice
    wmiEnumClassWin32MountPoint
    wmiEnumClassWin32MoveFileAction
    wmiEnumClassWin32MSIResource
    wmiEnumClassWin32NamedJobObject
    wmiEnumClassWin32NamedJobObjectActgInfo
    wmiEnumClassWin32NamedJobObjectLimit
    wmiEnumClassWin32NamedJobObjectLimitSetting
    wmiEnumClassWin32NamedJobObjectProcess
    wmiEnumClassWin32NamedJobObjectSecLimit
    wmiEnumClassWin32NamedJobObjectSecLimitSetting
    wmiEnumClassWin32NamedJobObjectStatistics
    wmiEnumClassWin32NetworkAdapter
    wmiEnumClassWin32NetworkAdapterConfiguration
    wmiEnumClassWin32NetworkAdapterSetting
    wmiEnumClassWin32NetworkClient
    wmiEnumClassWin32NetworkConnection
    wmiEnumClassWin32NetworkLoginProfile
    wmiEnumClassWin32NetworkProtocol
    wmiEnumClassWin32NTDomain
    wmiEnumClassWin32NTEventlogFile
    wmiEnumClassWin32NTLogEvent
    wmiEnumClassWin32NTLogEventComputer
    wmiEnumClassWin32NTLogEventLog
    wmiEnumClassWin32NTLogEventUser
    wmiEnumClassWin32ODBCAttribute
    wmiEnumClassWin32ODBCDataSourceAttribute
    wmiEnumClassWin32ODBCDataSourceSpecification
    wmiEnumClassWin32ODBCDriverAttribute
    wmiEnumClassWin32ODBCDriverSoftwareElement
    wmiEnumClassWin32ODBCDriverSpecification
    wmiEnumClassWin32ODBCSourceAttribute
    wmiEnumClassWin32ODBCTranslatorSpecification
    wmiEnumClassWin32OfflineFilesAssociatedItems
    wmiEnumClassWin32OfflineFilesBackgroundSync
    wmiEnumClassWin32OfflineFilesCache
    wmiEnumClassWin32OfflineFilesChangeInfo
    wmiEnumClassWin32OfflineFilesConnectionInfo
    wmiEnumClassWin32OfflineFilesDirtyInfo
    wmiEnumClassWin32OfflineFilesDiskSpaceLimit
    wmiEnumClassWin32OfflineFilesFileSysInfo
    wmiEnumClassWin32OfflineFilesHealth
    wmiEnumClassWin32OfflineFilesItem
    wmiEnumClassWin32OfflineFilesMachineConfiguration
    wmiEnumClassWin32OfflineFilesPinInfo
    wmiEnumClassWin32OfflineFilesSuspendInfo
    wmiEnumClassWin32OfflineFilesUserConfiguration
    wmiEnumClassWin32OnBoardDevice
    wmiEnumClassWin32OperatingSystem
    wmiEnumClassWin32OperatingSystemAutochkSetting
    wmiEnumClassWin32OperatingSystemQFE
    wmiEnumClassWin32OptionalFeature
    wmiEnumClassWin32OSRecoveryConfiguration
    wmiEnumClassWin32PageFile
    wmiEnumClassWin32PageFileElementSetting
    wmiEnumClassWin32PageFileSetting
    wmiEnumClassWin32PageFileUsage
    wmiEnumClassWin32ParallelPort
    wmiEnumClassWin32Patch
    wmiEnumClassWin32PatchFile
    wmiEnumClassWin32PatchPackage
    wmiEnumClassWin32PCMCIAController
    wmiEnumClassWin32PhysicalMedia
    wmiEnumClassWin32PhysicalMemory
    wmiEnumClassWin32PhysicalMemoryArray
    wmiEnumClassWin32PhysicalMemoryLocation
    wmiEnumClassWin32PingStatus
    wmiEnumClassWin32PNPAllocatedResource
    wmiEnumClassWin32PnPDevice
    wmiEnumClassWin32PnPEntity
    wmiEnumClassWin32PnPSignedDriver
    wmiEnumClassWin32PnPSignedDriverCIMDataFile
    wmiEnumClassWin32PointingDevice
    wmiEnumClassWin32PortableBattery
    wmiEnumClassWin32PortConnector
    wmiEnumClassWin32PortResource
    wmiEnumClassWin32POTSModem
    wmiEnumClassWin32POTSModemToSerialPort
    wmiEnumClassWin32PowerManagementEvent
    wmiEnumClassWin32Printer
    wmiEnumClassWin32PrinterConfiguration
    wmiEnumClassWin32PrinterController
    wmiEnumClassWin32PrinterDriver
    wmiEnumClassWin32PrinterDriverDll
    wmiEnumClassWin32PrinterSetting
    wmiEnumClassWin32PrinterShare
    wmiEnumClassWin32PrintJob
    wmiEnumClassWin32PrivilegesStatus
    wmiEnumClassWin32Process
    wmiEnumClassWin32Processor
    wmiEnumClassWin32ProcessStartTrace
    wmiEnumClassWin32ProcessStartup
    wmiEnumClassWin32ProcessStopTrace
    wmiEnumClassWin32ProcessTrace
    wmiEnumClassWin32Product
    wmiEnumClassWin32ProductCheck
    wmiEnumClassWin32ProductResource
    wmiEnumClassWin32ProductSoftwareFeatures
    wmiEnumClassWin32ProgIDSpecification
    wmiEnumClassWin32ProgramGroupContents
    wmiEnumClassWin32ProgramGroupOrItem
    wmiEnumClassWin32Property
    wmiEnumClassWin32ProtocolBinding
    wmiEnumClassWin32PublishComponentAction
    wmiEnumClassWin32QuickFixEngineering
    wmiEnumClassWin32QuotaSetting
    wmiEnumClassWin32Refrigeration
    wmiEnumClassWin32Registry
    wmiEnumClassWin32RegistryAction
    wmiEnumClassWin32Reliability
    wmiEnumClassWin32ReliabilityRecords
    wmiEnumClassWin32ReliabilityStabilityMetrics
    wmiEnumClassWin32RemoveFileAction
    wmiEnumClassWin32RemoveIniAction
    wmiEnumClassWin32ReserveCost
    wmiEnumClassWin32RoamingProfileBackgroundUploadParams
    wmiEnumClassWin32RoamingProfileMachineConfiguration
    wmiEnumClassWin32RoamingProfileSlowLinkParams
    wmiEnumClassWin32RoamingProfileUserConfiguration
    wmiEnumClassWin32RoamingUserHealthConfiguration
    wmiEnumClassWin32ScheduledJob
    wmiEnumClassWin32SCSIController
    wmiEnumClassWin32SCSIControllerDevice
    wmiEnumClassWin32SecurityDescriptor
    wmiEnumClassWin32SecurityDescriptorHelper
    wmiEnumClassWin32SecuritySetting
    wmiEnumClassWin32SecuritySettingAccess
    wmiEnumClassWin32SecuritySettingAuditing
    wmiEnumClassWin32SecuritySettingGroup
    wmiEnumClassWin32SecuritySettingOfLogicalFile
    wmiEnumClassWin32SecuritySettingOfLogicalShare
    wmiEnumClassWin32SecuritySettingOfObject
    wmiEnumClassWin32SecuritySettingOwner
    wmiEnumClassWin32SelfRegModuleAction
    wmiEnumClassWin32SerialPort
    wmiEnumClassWin32SerialPortConfiguration
    wmiEnumClassWin32SerialPortSetting
    wmiEnumClassWin32ServerConnection
    wmiEnumClassWin32ServerFeature
    wmiEnumClassWin32ServerSession
    wmiEnumClassWin32Service
    wmiEnumClassWin32ServiceControl
    wmiEnumClassWin32ServiceSpecification
    wmiEnumClassWin32ServiceSpecificationService
    wmiEnumClassWin32Session
    wmiEnumClassWin32SessionConnection
    wmiEnumClassWin32SessionProcess
    wmiEnumClassWin32SessionResource
    wmiEnumClassWin32SettingCheck
    wmiEnumClassWin32ShadowBy
    wmiEnumClassWin32ShadowContext
    wmiEnumClassWin32ShadowCopy
    wmiEnumClassWin32ShadowDiffVolumeSupport
    wmiEnumClassWin32ShadowFor
    wmiEnumClassWin32ShadowOn
    wmiEnumClassWin32ShadowProvider
    wmiEnumClassWin32ShadowStorage
    wmiEnumClassWin32ShadowVolumeSupport
    wmiEnumClassWin32Share
    wmiEnumClassWin32ShareToDirectory
    wmiEnumClassWin32ShortcutAction
    wmiEnumClassWin32ShortcutFile
    wmiEnumClassWin32ShortcutSAP
    wmiEnumClassWin32SID
    wmiEnumClassWin32SIDandAttributes
    wmiEnumClassWin32SMBIOSMemory
    wmiEnumClassWin32SoftwareElement
    wmiEnumClassWin32SoftwareElementAction
    wmiEnumClassWin32SoftwareElementCheck
    wmiEnumClassWin32SoftwareElementCondition
    wmiEnumClassWin32SoftwareElementResource
    wmiEnumClassWin32SoftwareFeature
    wmiEnumClassWin32SoftwareFeatureAction
    wmiEnumClassWin32SoftwareFeatureCheck
    wmiEnumClassWin32SoftwareFeatureParent
    wmiEnumClassWin32SoftwareFeatureSoftwareElements
    wmiEnumClassWin32SoundDevice
    wmiEnumClassWin32StartupCommand
    wmiEnumClassWin32SubDirectory
    wmiEnumClassWin32SubSession
    wmiEnumClassWin32SystemAccount
    wmiEnumClassWin32SystemBIOS
    wmiEnumClassWin32SystemBootConfiguration
    wmiEnumClassWin32SystemConfigurationChangeEvent
    wmiEnumClassWin32SystemDesktop
    wmiEnumClassWin32SystemDevices
    wmiEnumClassWin32SystemDriver
    wmiEnumClassWin32SystemDriverPNPEntity
    wmiEnumClassWin32SystemEnclosure
    wmiEnumClassWin32SystemLoadOrderGroups
    wmiEnumClassWin32SystemMemoryResource
    wmiEnumClassWin32SystemNetworkConnections
    wmiEnumClassWin32SystemOperatingSystem
    wmiEnumClassWin32SystemPartitions
    wmiEnumClassWin32SystemProcesses
    wmiEnumClassWin32SystemProgramGroups
    wmiEnumClassWin32SystemResources
    wmiEnumClassWin32SystemServices
    wmiEnumClassWin32SystemSetting
    wmiEnumClassWin32SystemSlot
    wmiEnumClassWin32SystemSystemDriver
    wmiEnumClassWin32SystemTimeZone
    wmiEnumClassWin32SystemTrace
    wmiEnumClassWin32SystemUsers
    wmiEnumClassWin32TapeDrive
    wmiEnumClassWin32TCPIPPrinterPort
    wmiEnumClassWin32TemperatureProbe
    wmiEnumClassWin32TerminalService
    wmiEnumClassWin32Thread
    wmiEnumClassWin32ThreadStartTrace
    wmiEnumClassWin32ThreadStopTrace
    wmiEnumClassWin32ThreadTrace
    wmiEnumClassWin32TimeZone
    wmiEnumClassWin32TokenGroups
    wmiEnumClassWin32TokenPrivileges
    wmiEnumClassWin32Trustee
    wmiEnumClassWin32TypeLibraryAction
    wmiEnumClassWin32USBController
    wmiEnumClassWin32USBControllerDevice
    wmiEnumClassWin32USBHub
    wmiEnumClassWin32UserAccount
    wmiEnumClassWin32UserDesktop
    wmiEnumClassWin32UserInDomain
    wmiEnumClassWin32UserProfile
    wmiEnumClassWin32UserStateConfigurationControls
    wmiEnumClassWin32UTCTime
    wmiEnumClassWin32VideoConfiguration
    wmiEnumClassWin32VideoController
    wmiEnumClassWin32VideoSettings
    wmiEnumClassWin32VoltageProbe
    wmiEnumClassWin32Volume
    wmiEnumClassWin32VolumeChangeEvent
    wmiEnumClassWin32VolumeQuota
    wmiEnumClassWin32VolumeQuotaSetting
    wmiEnumClassWin32VolumeUserQuota
    wmiEnumClassWin32WinSAT
    wmiEnumClassWin32WMIElementSetting
    wmiEnumClassWin32WMISetting
    wmiEnumClassWin32Perf
    wmiEnumClassWin32PerfFormattedData
    wmiEnumClassWin32PerfFormattedDataAFDCountersMicrosoftWinsockBSP
    wmiEnumClassWin32PerfFormattedDataAPPPOOLCountersProviderAPPPOOLWAS
    wmiEnumClassWin32PerfFormattedDataASPActiveServerPages
    wmiEnumClassWin32PerfFormattedDataASPNETASPNET
    wmiEnumClassWin32PerfFormattedDataASPNETASPNETApplications
    wmiEnumClassWin32PerfFormattedDataASPNET2050727ASPNETAppsv2050727
    wmiEnumClassWin32PerfFormattedDataASPNET2050727ASPNETv2050727
    wmiEnumClassWin32PerfFormattedDataASPNET4030319ASPNETAppsv4030319
    wmiEnumClassWin32PerfFormattedDataASPNET4030319ASPNETv4030319
    wmiEnumClassWin32PerfFormattedDataaspnetstateASPNETStateService
    wmiEnumClassWin32PerfFormattedDataAuthorizationManagerAuthorizationManagerApplications
    wmiEnumClassWin32PerfFormattedDataBalancerStatsHyperVDynamicMemoryBalancer
    wmiEnumClassWin32PerfFormattedDataBalancerStatsHyperVDynamicMemoryVM
    wmiEnumClassWin32PerfFormattedDataBITSBITSNetUtilization
    wmiEnumClassWin32PerfFormattedDataCountersDNS64Global
    wmiEnumClassWin32PerfFormattedDataCountersEventTracingforWindows
    wmiEnumClassWin32PerfFormattedDataCountersEventTracingforWindowsSession
    wmiEnumClassWin32PerfFormattedDataCountersFileSystemDiskActivity
    wmiEnumClassWin32PerfFormattedDataCountersGenericIKEv1AuthIPandIKEv2
    wmiEnumClassWin32PerfFormattedDataCountersHTTPService
    wmiEnumClassWin32PerfFormattedDataCountersHTTPServiceRequestQueues
    wmiEnumClassWin32PerfFormattedDataCountersHTTPServiceUrlGroups
    wmiEnumClassWin32PerfFormattedDataCountersHyperVDynamicMemoryIntegrationService
    wmiEnumClassWin32PerfFormattedDataCountersHyperVVirtualMachineBusPipes
    wmiEnumClassWin32PerfFormattedDataCountersIPHTTPSGlobal
    wmiEnumClassWin32PerfFormattedDataCountersIPHTTPSSession
    wmiEnumClassWin32PerfFormattedDataCountersIPsecAuthIPIPv4
    wmiEnumClassWin32PerfFormattedDataCountersIPsecAuthIPIPv6
    wmiEnumClassWin32PerfFormattedDataCountersIPsecConnections
    wmiEnumClassWin32PerfFormattedDataCountersIPsecDoSProtection
    wmiEnumClassWin32PerfFormattedDataCountersIPsecDriver
    wmiEnumClassWin32PerfFormattedDataCountersIPsecIKEv1IPv4
    wmiEnumClassWin32PerfFormattedDataCountersIPsecIKEv1IPv6
    wmiEnumClassWin32PerfFormattedDataCountersIPsecIKEv2IPv4
    wmiEnumClassWin32PerfFormattedDataCountersIPsecIKEv2IPv6
    wmiEnumClassWin32PerfFormattedDataCountersNetlogon
    wmiEnumClassWin32PerfFormattedDataCountersNetworkQoSPolicy
    wmiEnumClassWin32PerfFormattedDataCountersPacerFlow
    wmiEnumClassWin32PerfFormattedDataCountersPacerPipe
    wmiEnumClassWin32PerfFormattedDataCountersPacketDirectECUtilization
    wmiEnumClassWin32PerfFormattedDataCountersPacketDirectQueueDepth
    wmiEnumClassWin32PerfFormattedDataCountersPacketDirectReceiveCounters
    wmiEnumClassWin32PerfFormattedDataCountersPacketDirectReceiveFilters
    wmiEnumClassWin32PerfFormattedDataCountersPacketDirectTransmitCounters
    wmiEnumClassWin32PerfFormattedDataCountersPerProcessorNetworkActivityCycles
    wmiEnumClassWin32PerfFormattedDataCountersPerProcessorNetworkInterfaceCardActivity
    wmiEnumClassWin32PerfFormattedDataCountersPhysicalNetworkInterfaceCardActivity
    wmiEnumClassWin32PerfFormattedDataCountersPowerShellWorkflow
    wmiEnumClassWin32PerfFormattedDataCountersProcessorInformation
    wmiEnumClassWin32PerfFormattedDataCountersRDMAActivity
    wmiEnumClassWin32PerfFormattedDataCountersRemoteFXGraphics
    wmiEnumClassWin32PerfFormattedDataCountersRemoteFXNetwork
    wmiEnumClassWin32PerfFormattedDataCountersSMBClientShares
    wmiEnumClassWin32PerfFormattedDataCountersSMBServer
    wmiEnumClassWin32PerfFormattedDataCountersSMBServerSessions
    wmiEnumClassWin32PerfFormattedDataCountersSMBServerShares
    wmiEnumClassWin32PerfFormattedDataCountersStorageSpacesTier
    wmiEnumClassWin32PerfFormattedDataCountersStorageSpacesWriteCache
    wmiEnumClassWin32PerfFormattedDataCountersSynchronization
    wmiEnumClassWin32PerfFormattedDataCountersSynchronizationNuma
    wmiEnumClassWin32PerfFormattedDataCountersTeredoClient
    wmiEnumClassWin32PerfFormattedDataCountersTeredoRelay
    wmiEnumClassWin32PerfFormattedDataCountersTeredoServer
    wmiEnumClassWin32PerfFormattedDataCountersThermalZoneInformation
    wmiEnumClassWin32PerfFormattedDataCountersWFP
    wmiEnumClassWin32PerfFormattedDataCountersWFPv4
    wmiEnumClassWin32PerfFormattedDataCountersWFPv6
    wmiEnumClassWin32PerfFormattedDataCountersWSManQuotaStatistics
    wmiEnumClassWin32PerfFormattedDataCountersXHCICommonBuffer
    wmiEnumClassWin32PerfFormattedDataCountersXHCIInterrupter
    wmiEnumClassWin32PerfFormattedDataCountersXHCITransferRing
    wmiEnumClassWin32PerfFormattedDataDdmCounterProviderRAS
    wmiEnumClassWin32PerfFormattedDataDeliveryOptimizationDeliveryOptimizationSwarm
    wmiEnumClassWin32PerfFormattedDataDistributedRoutingTablePerfDistributedRoutingTable
    wmiEnumClassWin32PerfFormattedDataESENTDatabase
    wmiEnumClassWin32PerfFormattedDataESENTDatabaseInstances
    wmiEnumClassWin32PerfFormattedDataESENTDatabaseTableClasses
    wmiEnumClassWin32PerfFormattedDataEthernetPerfProviderHyperVLegacyNetworkAdapter
    wmiEnumClassWin32PerfFormattedDataFaxServiceFaxService
    wmiEnumClassWin32PerfFormattedDataftpsvcMicrosoftFTPService
    wmiEnumClassWin32PerfFormattedDataGmoPerfProviderHyperVVMSaveSnapshotandRestore
    wmiEnumClassWin32PerfFormattedDataHvStatsHyperVHypervisor
    wmiEnumClassWin32PerfFormattedDataHvStatsHyperVHypervisorLogicalProcessor
    wmiEnumClassWin32PerfFormattedDataHvStatsHyperVHypervisorPartition
    wmiEnumClassWin32PerfFormattedDataHvStatsHyperVHypervisorRootPartition
    wmiEnumClassWin32PerfFormattedDataHvStatsHyperVHypervisorRootVirtualProcessor
    wmiEnumClassWin32PerfFormattedDataHvStatsHyperVHypervisorVirtualProcessor
    wmiEnumClassWin32PerfFormattedDataIdePerfProviderHyperVVirtualIDEController
    wmiEnumClassWin32PerfFormattedDataLocalSessionManagerTerminalServices
    wmiEnumClassWin32PerfFormattedDataLsaSecurityPerProcessStatistics
    wmiEnumClassWin32PerfFormattedDataLsaSecuritySystemWideStatistics
    wmiEnumClassWin32PerfFormattedDataMicrosoftWindowsBitLockerDriverCountersProviderBitLocker
    wmiEnumClassWin32PerfFormattedDataMicrosoftWindowsRemoteDesktopServicesRemoteFXSynth3dvscRemoteFXSynth3DVSCVMDevice
    wmiEnumClassWin32PerfFormattedDataMicrosoftWindowsRemoteDesktopServicesRemoteFXSynth3dvscRemoteFXSynth3DVSCVMTransportChannel
    wmiEnumClassWin32PerfFormattedDataMSDTCDistributedTransactionCoordinator
    wmiEnumClassWin32PerfFormattedDataMSDTCBridge3000MSDTCBridge3000
    wmiEnumClassWin32PerfFormattedDataMSDTCBridge4000MSDTCBridge4000
    wmiEnumClassWin32PerfFormattedDataNETCLRDataNETCLRData
    wmiEnumClassWin32PerfFormattedDataNETCLRNetworkingNETCLRNetworking
    wmiEnumClassWin32PerfFormattedDataNETCLRNetworking4000NETCLRNetworking4000
    wmiEnumClassWin32PerfFormattedDataNETDataProviderforOracleNETDataProviderforOracle
    wmiEnumClassWin32PerfFormattedDataNETDataProviderforSqlServerNETDataProviderforSqlServer
    wmiEnumClassWin32PerfFormattedDataNETFrameworkNETCLRExceptions
    wmiEnumClassWin32PerfFormattedDataNETFrameworkNETCLRInterop
    wmiEnumClassWin32PerfFormattedDataNETFrameworkNETCLRJit
    wmiEnumClassWin32PerfFormattedDataNETFrameworkNETCLRLoading
    wmiEnumClassWin32PerfFormattedDataNETFrameworkNETCLRLocksAndThreads
    wmiEnumClassWin32PerfFormattedDataNETFrameworkNETCLRMemory
    wmiEnumClassWin32PerfFormattedDataNETFrameworkNETCLRRemoting
    wmiEnumClassWin32PerfFormattedDataNETFrameworkNETCLRSecurity
    wmiEnumClassWin32PerfFormattedDataNETMemoryCache40NETMemoryCache40
    wmiEnumClassWin32PerfFormattedDataNvspNicStatsHyperVVirtualNetworkAdapter
    wmiEnumClassWin32PerfFormattedDataNvspPortStatsHyperVVirtualSwitchPort
    wmiEnumClassWin32PerfFormattedDataNvspSwitchStatsHyperVVirtualSwitch
    wmiEnumClassWin32PerfFormattedDataOfflineFilesClientSideCaching
    wmiEnumClassWin32PerfFormattedDataOfflineFilesOfflineFiles
    wmiEnumClassWin32PerfFormattedDataPeerDistSvcBranchCache
    wmiEnumClassWin32PerfFormattedDataPeerNameResolutionProtocolPerfPeerNameResolutionProtocol
    wmiEnumClassWin32PerfFormattedDataPerfDiskLogicalDisk
    wmiEnumClassWin32PerfFormattedDataPerfDiskPhysicalDisk
    wmiEnumClassWin32PerfFormattedDataPerfNetBrowser
    wmiEnumClassWin32PerfFormattedDataPerfNetRedirector
    wmiEnumClassWin32PerfFormattedDataPerfNetServer
    wmiEnumClassWin32PerfFormattedDataPerfNetServerWorkQueues
    wmiEnumClassWin32PerfFormattedDataPerfOSCache
    wmiEnumClassWin32PerfFormattedDataPerfOSMemory
    wmiEnumClassWin32PerfFormattedDataPerfOSNUMANodeMemory
    wmiEnumClassWin32PerfFormattedDataPerfOSObjects
    wmiEnumClassWin32PerfFormattedDataPerfOSPagingFile
    wmiEnumClassWin32PerfFormattedDataPerfOSProcessor
    wmiEnumClassWin32PerfFormattedDataPerfOSSystem
    wmiEnumClassWin32PerfFormattedDataPerfProcFullImageCostly
    wmiEnumClassWin32PerfFormattedDataPerfProcImageCostly
    wmiEnumClassWin32PerfFormattedDataPerfProcJobObject
    wmiEnumClassWin32PerfFormattedDataPerfProcJobObjectDetails
    wmiEnumClassWin32PerfFormattedDataPerfProcProcess
    wmiEnumClassWin32PerfFormattedDataPerfProcProcessAddressSpaceCostly
    wmiEnumClassWin32PerfFormattedDataPerfProcThread
    wmiEnumClassWin32PerfFormattedDataPerfProcThreadDetailsCostly
    wmiEnumClassWin32PerfFormattedDataPowerMeterCounterEnergyMeter
    wmiEnumClassWin32PerfFormattedDataPowerMeterCounterPowerMeter
    wmiEnumClassWin32PerfFormattedDatardyboostReadyBoostCache
    wmiEnumClassWin32PerfFormattedDataRemoteAccessRASPort
    wmiEnumClassWin32PerfFormattedDataRemoteAccessRASTotal
    wmiEnumClassWin32PerfFormattedDataRemotePerfProviderHyperVVMRemoting
    wmiEnumClassWin32PerfFormattedDataServiceModel4000ServiceModelEndpoint4000
    wmiEnumClassWin32PerfFormattedDataServiceModel4000ServiceModelOperation4000
    wmiEnumClassWin32PerfFormattedDataServiceModel4000ServiceModelService4000
    wmiEnumClassWin32PerfFormattedDataServiceModelEndpoint3000ServiceModelEndpoint3000
    wmiEnumClassWin32PerfFormattedDataServiceModelOperation3000ServiceModelOperation3000
    wmiEnumClassWin32PerfFormattedDataServiceModelService3000ServiceModelService3000
    wmiEnumClassWin32PerfFormattedDataSMSvcHost3000SMSvcHost3000
    wmiEnumClassWin32PerfFormattedDataSMSvcHost4000SMSvcHost4000
    wmiEnumClassWin32PerfFormattedDataSpoolerPrintQueue
    wmiEnumClassWin32PerfFormattedDataStorageStatsHyperVVirtualStorageDevice
    wmiEnumClassWin32PerfFormattedDataTapiSrvTelephony
    wmiEnumClassWin32PerfFormattedDataTBSTBScounters
    wmiEnumClassWin32PerfFormattedDataTcpipICMP
    wmiEnumClassWin32PerfFormattedDataTcpipICMPv6
    wmiEnumClassWin32PerfFormattedDataTcpipIPv4
    wmiEnumClassWin32PerfFormattedDataTcpipIPv6
    wmiEnumClassWin32PerfFormattedDataTcpipNBTConnection
    wmiEnumClassWin32PerfFormattedDataTcpipNetworkAdapter
    wmiEnumClassWin32PerfFormattedDataTcpipNetworkInterface
    wmiEnumClassWin32PerfFormattedDataTcpipTCPv4
    wmiEnumClassWin32PerfFormattedDataTcpipTCPv6
    wmiEnumClassWin32PerfFormattedDataTcpipUDPv4
    wmiEnumClassWin32PerfFormattedDataTcpipUDPv6
    wmiEnumClassWin32PerfFormattedDataTCPIPCountersTCPIPPerformanceDiagnostics
    wmiEnumClassWin32PerfFormattedDataTermServiceTerminalServicesSession
    wmiEnumClassWin32PerfFormattedDataUGathererSearchGathererProjects
    wmiEnumClassWin32PerfFormattedDataUGTHRSVCSearchGatherer
    wmiEnumClassWin32PerfFormattedDatausbhubUSB
    wmiEnumClassWin32PerfFormattedDataVidPerfProviderHyperVVMVidNumaNode
    wmiEnumClassWin32PerfFormattedDataVidPerfProviderHyperVVMVidPartition
    wmiEnumClassWin32PerfFormattedDataVmbusStatsHyperVVirtualMachineBus
    wmiEnumClassWin32PerfFormattedDataVmmsVirtualMachineStatsHyperVVirtualMachineHealthSummary
    wmiEnumClassWin32PerfFormattedDataVmmsVirtualMachineStatsHyperVVirtualMachineSummary
    wmiEnumClassWin32PerfFormattedDataVmTaskManagerStatsHyperVTaskManagerDetail
    wmiEnumClassWin32PerfFormattedDataW3SVCWebService
    wmiEnumClassWin32PerfFormattedDataW3SVCWebServiceCache
    wmiEnumClassWin32PerfFormattedDataW3SVCW3WPCounterProviderW3SVCW3WP
    wmiEnumClassWin32PerfFormattedDataWASW3WPCounterProviderWASW3WP
    wmiEnumClassWin32PerfFormattedDataWindowsMediaPlayerWindowsMediaPlayerMetadata
    wmiEnumClassWin32PerfFormattedDataWindowsWorkflowFoundation3000WindowsWorkflowFoundation
    wmiEnumClassWin32PerfFormattedDataWindowsWorkflowFoundation4000WFSystemWorkflow4000
    wmiEnumClassWin32PerfFormattedDataWorkflowServiceHost4000WorkflowServiceHost4000
    wmiEnumClassWin32PerfFormattedDataWSearchIdxPiSearchIndexer
    wmiEnumClassWin32PerfRawData
    wmiEnumClassWin32PerfRawDataAFDCountersMicrosoftWinsockBSP
    wmiEnumClassWin32PerfRawDataAPPPOOLCountersProviderAPPPOOLWAS
    wmiEnumClassWin32PerfRawDataASPActiveServerPages
    wmiEnumClassWin32PerfRawDataASPNETASPNET
    wmiEnumClassWin32PerfRawDataASPNETASPNETApplications
    wmiEnumClassWin32PerfRawDataASPNET2050727ASPNETAppsv2050727
    wmiEnumClassWin32PerfRawDataASPNET2050727ASPNETv2050727
    wmiEnumClassWin32PerfRawDataASPNET4030319ASPNETAppsv4030319
    wmiEnumClassWin32PerfRawDataASPNET4030319ASPNETv4030319
    wmiEnumClassWin32PerfRawDataaspnetstateASPNETStateService
    wmiEnumClassWin32PerfRawDataAuthorizationManagerAuthorizationManagerApplications
    wmiEnumClassWin32PerfRawDataBalancerStatsHyperVDynamicMemoryBalancer
    wmiEnumClassWin32PerfRawDataBalancerStatsHyperVDynamicMemoryVM
    wmiEnumClassWin32PerfRawDataBITSBITSNetUtilization
    wmiEnumClassWin32PerfRawDataCountersDNS64Global
    wmiEnumClassWin32PerfRawDataCountersEventTracingforWindows
    wmiEnumClassWin32PerfRawDataCountersEventTracingforWindowsSession
    wmiEnumClassWin32PerfRawDataCountersFileSystemDiskActivity
    wmiEnumClassWin32PerfRawDataCountersGenericIKEv1AuthIPandIKEv2
    wmiEnumClassWin32PerfRawDataCountersHTTPService
    wmiEnumClassWin32PerfRawDataCountersHTTPServiceRequestQueues
    wmiEnumClassWin32PerfRawDataCountersHTTPServiceUrlGroups
    wmiEnumClassWin32PerfRawDataCountersHyperVDynamicMemoryIntegrationService
    wmiEnumClassWin32PerfRawDataCountersHyperVVirtualMachineBusPipes
    wmiEnumClassWin32PerfRawDataCountersIPHTTPSGlobal
    wmiEnumClassWin32PerfRawDataCountersIPHTTPSSession
    wmiEnumClassWin32PerfRawDataCountersIPsecAuthIPIPv4
    wmiEnumClassWin32PerfRawDataCountersIPsecAuthIPIPv6
    wmiEnumClassWin32PerfRawDataCountersIPsecConnections
    wmiEnumClassWin32PerfRawDataCountersIPsecDoSProtection
    wmiEnumClassWin32PerfRawDataCountersIPsecDriver
    wmiEnumClassWin32PerfRawDataCountersIPsecIKEv1IPv4
    wmiEnumClassWin32PerfRawDataCountersIPsecIKEv1IPv6
    wmiEnumClassWin32PerfRawDataCountersIPsecIKEv2IPv4
    wmiEnumClassWin32PerfRawDataCountersIPsecIKEv2IPv6
    wmiEnumClassWin32PerfRawDataCountersNetlogon
    wmiEnumClassWin32PerfRawDataCountersNetworkQoSPolicy
    wmiEnumClassWin32PerfRawDataCountersPacerFlow
    wmiEnumClassWin32PerfRawDataCountersPacerPipe
    wmiEnumClassWin32PerfRawDataCountersPacketDirectECUtilization
    wmiEnumClassWin32PerfRawDataCountersPacketDirectQueueDepth
    wmiEnumClassWin32PerfRawDataCountersPacketDirectReceiveCounters
    wmiEnumClassWin32PerfRawDataCountersPacketDirectReceiveFilters
    wmiEnumClassWin32PerfRawDataCountersPacketDirectTransmitCounters
    wmiEnumClassWin32PerfRawDataCountersPerProcessorNetworkActivityCycles
    wmiEnumClassWin32PerfRawDataCountersPerProcessorNetworkInterfaceCardActivity
    wmiEnumClassWin32PerfRawDataCountersPhysicalNetworkInterfaceCardActivity
    wmiEnumClassWin32PerfRawDataCountersPowerShellWorkflow
    wmiEnumClassWin32PerfRawDataCountersProcessorInformation
    wmiEnumClassWin32PerfRawDataCountersRDMAActivity
    wmiEnumClassWin32PerfRawDataCountersRemoteFXGraphics
    wmiEnumClassWin32PerfRawDataCountersRemoteFXNetwork
    wmiEnumClassWin32PerfRawDataCountersSMBClientShares
    wmiEnumClassWin32PerfRawDataCountersSMBServer
    wmiEnumClassWin32PerfRawDataCountersSMBServerSessions
    wmiEnumClassWin32PerfRawDataCountersSMBServerShares
    wmiEnumClassWin32PerfRawDataCountersStorageSpacesTier
    wmiEnumClassWin32PerfRawDataCountersStorageSpacesWriteCache
    wmiEnumClassWin32PerfRawDataCountersSynchronization
    wmiEnumClassWin32PerfRawDataCountersSynchronizationNuma
    wmiEnumClassWin32PerfRawDataCountersTeredoClient
    wmiEnumClassWin32PerfRawDataCountersTeredoRelay
    wmiEnumClassWin32PerfRawDataCountersTeredoServer
    wmiEnumClassWin32PerfRawDataCountersThermalZoneInformation
    wmiEnumClassWin32PerfRawDataCountersWFP
    wmiEnumClassWin32PerfRawDataCountersWFPv4
    wmiEnumClassWin32PerfRawDataCountersWFPv6
    wmiEnumClassWin32PerfRawDataCountersWSManQuotaStatistics
    wmiEnumClassWin32PerfRawDataCountersXHCICommonBuffer
    wmiEnumClassWin32PerfRawDataCountersXHCIInterrupter
    wmiEnumClassWin32PerfRawDataCountersXHCITransferRing
    wmiEnumClassWin32PerfRawDataDdmCounterProviderRAS
    wmiEnumClassWin32PerfRawDataDeliveryOptimizationDeliveryOptimizationSwarm
    wmiEnumClassWin32PerfRawDataDistributedRoutingTablePerfDistributedRoutingTable
    wmiEnumClassWin32PerfRawDataESENTDatabase
    wmiEnumClassWin32PerfRawDataESENTDatabaseInstances
    wmiEnumClassWin32PerfRawDataESENTDatabaseTableClasses
    wmiEnumClassWin32PerfRawDataEthernetPerfProviderHyperVLegacyNetworkAdapter
    wmiEnumClassWin32PerfRawDataFaxServiceFaxService
    wmiEnumClassWin32PerfRawDataftpsvcMicrosoftFTPService
    wmiEnumClassWin32PerfRawDataGmoPerfProviderHyperVVMSaveSnapshotandRestore
    wmiEnumClassWin32PerfRawDataHvStatsHyperVHypervisor
    wmiEnumClassWin32PerfRawDataHvStatsHyperVHypervisorLogicalProcessor
    wmiEnumClassWin32PerfRawDataHvStatsHyperVHypervisorPartition
    wmiEnumClassWin32PerfRawDataHvStatsHyperVHypervisorRootPartition
    wmiEnumClassWin32PerfRawDataHvStatsHyperVHypervisorRootVirtualProcessor
    wmiEnumClassWin32PerfRawDataHvStatsHyperVHypervisorVirtualProcessor
    wmiEnumClassWin32PerfRawDataIdePerfProviderHyperVVirtualIDEController
    wmiEnumClassWin32PerfRawDataLocalSessionManagerTerminalServices
    wmiEnumClassWin32PerfRawDataLsaSecurityPerProcessStatistics
    wmiEnumClassWin32PerfRawDataLsaSecuritySystemWideStatistics
    wmiEnumClassWin32PerfRawDataMicrosoftWindowsBitLockerDriverCountersProviderBitLocker
    wmiEnumClassWin32PerfRawDataMicrosoftWindowsRemoteDesktopServicesRemoteFXSynth3dvscRemoteFXSynth3DVSCVMDevice
    wmiEnumClassWin32PerfRawDataMicrosoftWindowsRemoteDesktopServicesRemoteFXSynth3dvscRemoteFXSynth3DVSCVMTransportChannel
    wmiEnumClassWin32PerfRawDataMSDTCDistributedTransactionCoordinator
    wmiEnumClassWin32PerfRawDataMSDTCBridge3000MSDTCBridge3000
    wmiEnumClassWin32PerfRawDataMSDTCBridge4000MSDTCBridge4000
    wmiEnumClassWin32PerfRawDataNETCLRDataNETCLRData
    wmiEnumClassWin32PerfRawDataNETCLRNetworkingNETCLRNetworking
    wmiEnumClassWin32PerfRawDataNETCLRNetworking4000NETCLRNetworking4000
    wmiEnumClassWin32PerfRawDataNETDataProviderforOracleNETDataProviderforOracle
    wmiEnumClassWin32PerfRawDataNETDataProviderforSqlServerNETDataProviderforSqlServer
    wmiEnumClassWin32PerfRawDataNETFrameworkNETCLRExceptions
    wmiEnumClassWin32PerfRawDataNETFrameworkNETCLRInterop
    wmiEnumClassWin32PerfRawDataNETFrameworkNETCLRJit
    wmiEnumClassWin32PerfRawDataNETFrameworkNETCLRLoading
    wmiEnumClassWin32PerfRawDataNETFrameworkNETCLRLocksAndThreads
    wmiEnumClassWin32PerfRawDataNETFrameworkNETCLRMemory
    wmiEnumClassWin32PerfRawDataNETFrameworkNETCLRRemoting
    wmiEnumClassWin32PerfRawDataNETFrameworkNETCLRSecurity
    wmiEnumClassWin32PerfRawDataNETMemoryCache40NETMemoryCache40
    wmiEnumClassWin32PerfRawDataNvspNicStatsHyperVVirtualNetworkAdapter
    wmiEnumClassWin32PerfRawDataNvspPortStatsHyperVVirtualSwitchPort
    wmiEnumClassWin32PerfRawDataNvspSwitchStatsHyperVVirtualSwitch
    wmiEnumClassWin32PerfRawDataOfflineFilesClientSideCaching
    wmiEnumClassWin32PerfRawDataOfflineFilesOfflineFiles
    wmiEnumClassWin32PerfRawDataPeerDistSvcBranchCache
    wmiEnumClassWin32PerfRawDataPeerNameResolutionProtocolPerfPeerNameResolutionProtocol
    wmiEnumClassWin32PerfRawDataPerfDiskLogicalDisk
    wmiEnumClassWin32PerfRawDataPerfDiskPhysicalDisk
    wmiEnumClassWin32PerfRawDataPerfNetBrowser
    wmiEnumClassWin32PerfRawDataPerfNetRedirector
    wmiEnumClassWin32PerfRawDataPerfNetServer
    wmiEnumClassWin32PerfRawDataPerfNetServerWorkQueues
    wmiEnumClassWin32PerfRawDataPerfOSCache
    wmiEnumClassWin32PerfRawDataPerfOSMemory
    wmiEnumClassWin32PerfRawDataPerfOSNUMANodeMemory
    wmiEnumClassWin32PerfRawDataPerfOSObjects
    wmiEnumClassWin32PerfRawDataPerfOSPagingFile
    wmiEnumClassWin32PerfRawDataPerfOSProcessor
    wmiEnumClassWin32PerfRawDataPerfOSSystem
    wmiEnumClassWin32PerfRawDataPerfProcFullImageCostly
    wmiEnumClassWin32PerfRawDataPerfProcImageCostly
    wmiEnumClassWin32PerfRawDataPerfProcJobObject
    wmiEnumClassWin32PerfRawDataPerfProcJobObjectDetails
    wmiEnumClassWin32PerfRawDataPerfProcProcess
    wmiEnumClassWin32PerfRawDataPerfProcProcessAddressSpaceCostly
    wmiEnumClassWin32PerfRawDataPerfProcThread
    wmiEnumClassWin32PerfRawDataPerfProcThreadDetailsCostly
    wmiEnumClassWin32PerfRawDataPowerMeterCounterEnergyMeter
    wmiEnumClassWin32PerfRawDataPowerMeterCounterPowerMeter
    wmiEnumClassWin32PerfRawDatardyboostReadyBoostCache
    wmiEnumClassWin32PerfRawDataRemoteAccessRASPort
    wmiEnumClassWin32PerfRawDataRemoteAccessRASTotal
    wmiEnumClassWin32PerfRawDataRemotePerfProviderHyperVVMRemoting
    wmiEnumClassWin32PerfRawDataServiceModel4000ServiceModelEndpoint4000
    wmiEnumClassWin32PerfRawDataServiceModel4000ServiceModelOperation4000
    wmiEnumClassWin32PerfRawDataServiceModel4000ServiceModelService4000
    wmiEnumClassWin32PerfRawDataServiceModelEndpoint3000ServiceModelEndpoint3000
    wmiEnumClassWin32PerfRawDataServiceModelOperation3000ServiceModelOperation3000
    wmiEnumClassWin32PerfRawDataServiceModelService3000ServiceModelService3000
    wmiEnumClassWin32PerfRawDataSMSvcHost3000SMSvcHost3000
    wmiEnumClassWin32PerfRawDataSMSvcHost4000SMSvcHost4000
    wmiEnumClassWin32PerfRawDataSpoolerPrintQueue
    wmiEnumClassWin32PerfRawDataStorageStatsHyperVVirtualStorageDevice
    wmiEnumClassWin32PerfRawDataTapiSrvTelephony
    wmiEnumClassWin32PerfRawDataTBSTBScounters
    wmiEnumClassWin32PerfRawDataTcpipICMP
    wmiEnumClassWin32PerfRawDataTcpipICMPv6
    wmiEnumClassWin32PerfRawDataTcpipIPv4
    wmiEnumClassWin32PerfRawDataTcpipIPv6
    wmiEnumClassWin32PerfRawDataTcpipNBTConnection
    wmiEnumClassWin32PerfRawDataTcpipNetworkAdapter
    wmiEnumClassWin32PerfRawDataTcpipNetworkInterface
    wmiEnumClassWin32PerfRawDataTcpipTCPv4
    wmiEnumClassWin32PerfRawDataTcpipTCPv6
    wmiEnumClassWin32PerfRawDataTcpipUDPv4
    wmiEnumClassWin32PerfRawDataTcpipUDPv6
    wmiEnumClassWin32PerfRawDataTCPIPCountersTCPIPPerformanceDiagnostics
    wmiEnumClassWin32PerfRawDataTermServiceTerminalServicesSession
    wmiEnumClassWin32PerfRawDataUGathererSearchGathererProjects
    wmiEnumClassWin32PerfRawDataUGTHRSVCSearchGatherer
    wmiEnumClassWin32PerfRawDatausbhubUSB
    wmiEnumClassWin32PerfRawDataVidPerfProviderHyperVVMVidNumaNode
    wmiEnumClassWin32PerfRawDataVidPerfProviderHyperVVMVidPartition
    wmiEnumClassWin32PerfRawDataVmbusStatsHyperVVirtualMachineBus
    wmiEnumClassWin32PerfRawDataVmmsVirtualMachineStatsHyperVVirtualMachineHealthSummary
    wmiEnumClassWin32PerfRawDataVmmsVirtualMachineStatsHyperVVirtualMachineSummary
    wmiEnumClassWin32PerfRawDataVmTaskManagerStatsHyperVTaskManagerDetail
    wmiEnumClassWin32PerfRawDataW3SVCWebService
    wmiEnumClassWin32PerfRawDataW3SVCWebServiceCache
    wmiEnumClassWin32PerfRawDataW3SVCW3WPCounterProviderW3SVCW3WP
    wmiEnumClassWin32PerfRawDataWASW3WPCounterProviderWASW3WP
    wmiEnumClassWin32PerfRawDataWindowsMediaPlayerWindowsMediaPlayerMetadata
    wmiEnumClassWin32PerfRawDataWindowsWorkflowFoundation3000WindowsWorkflowFoundation
    wmiEnumClassWin32PerfRawDataWindowsWorkflowFoundation4000WFSystemWorkflow4000
    wmiEnumClassWin32PerfRawDataWorkflowServiceHost4000WorkflowServiceHost4000
    wmiEnumClassWin32PerfRawDataWSearchIdxPiSearchIndexer
    wmiEnumClassEventViewerConsumer
    wmiEnumClassNTEventlogProviderConfig
    wmiEnumClassOfficeSoftwareProtectionProduct
    wmiEnumClassOfficeSoftwareProtectionService
    wmiEnumClassOfficeSoftwareProtectionTokenActivationLicense
    wmiEnumClassRegistryEvent
    wmiEnumClassRegistryKeyChangeEvent
    wmiEnumClassRegistryTreeChangeEvent
    wmiEnumClassRegistryValueChangeEvent
    wmiEnumClassScriptingStandardConsumerSetting
    wmiEnumClassSoftwareLicensingProduct
    wmiEnumClassSoftwareLicensingService
    wmiEnumClassSoftwareLicensingTokenActivationLicense
    wmiEnumClassStdRegProv
End Enum

'******************************************************************************
'* 定数定義
'******************************************************************************
'*-----------------------------------------------------------------------------
'* 拡張定数定義
'*-----------------------------------------------------------------------------
'*-----------------------------------------------------------------------------
'* Win32_Process クラス
'* Win32 システムのイベントのシーケンスを表します。
'* プロセッサまたはインタープリター、実行可能コードの一部、入力セットの一連の
'* 連続する操作は、このクラスの子孫 (またはメンバー) です。
'* WMI Provider : CIMWin32
'* UUID : {8502C4DC-5FBB-11D2-AAC1-006008C78BC7}
'*
'* @see http://www.wmifun.net/library/win32_process.html
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32Process = "Win32_Process"

'*-----------------------------------------------------------------------------
'* Win32_NetworkAdapterConfiguration クラス
'* ネットワーク アダプターの属性と動作を表します。
'* このクラスは拡張され、TCP/IP プロトコルの管理がサポートされる (ネットワーク
'*  アダプターから独立している) 追加のプロパティとメソッドが含まれます。
'* WMI Provider : CIMWin32
'* UUID : {8502C515-5FBB-11D2-AAC1-006008C78BC7}
'*
'* @see http://www.wmifun.net/library/win32_networkadapterconfiguration.html
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32NetworkAdapterConfiguration = "Win32_NetworkAdapterConfiguration"

'*-----------------------------------------------------------------------------
'* Win32_NTLogEvent クラス
'* このクラスは NT Eventlog からのインスタンスを変換するのに使用されます。
'* WMI Provider : MS_NT_EVENTLOG_PROVIDER
'* UUID : {8502C57C-5FBB-11D2-AAC1-006008C78BC7}
'*
'* @see http://www.wmifun.net/library/win32_ntlogevent.html
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32NTLogEvent = "Win32_NTLogEvent"

'*-----------------------------------------------------------------------------
'* Win32_OperatingSystem クラス
'* Win32 コンピューター システムにインストールされているオペレーティング
'* システムを表します。
'* Win32 システムにインストールされているオペレーティング システムはこのクラス
'* の子孫 (またはメンバー) です。
'* WMI Provider : CIMWin32
'* UUID : {8502C4DE-5FBB-11D2-AAC1-006008C78BC7}
'*
'* @see http://www.wmifun.net/library/win32_operatingsystem.html
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32OperatingSystem = "Win32_OperatingSystem"

'*-----------------------------------------------------------------------------
'* Win32_Printer クラス
'* プリンターの LogicalDevice の機能と管理です。
'* WMI Provider : CIMWin32
'* UUID : {8502C4BC-5FBB-11D2-AAC1-006008C78BC7}
'*
'* @see http://www.wmifun.net/library/win32_printer.html
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32Printer = "Win32_Printer"

'*-----------------------------------------------------------------------------
'* Win32_ComputerSystem クラス
'* Win32 環境で作動するコンピューター システムを表します。
'* WMI Provider : CIMWin32
'* UUID : {8502C4B0-5FBB-11D2-AAC1-006008C78BC7}
'*
'* @see http://www.wmifun.net/library/win32_computersystem.html
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ComputerSystem = "Win32_ComputerSystem"

'*-----------------------------------------------------------------------------
'* Win32_Processor クラス
'* Win32 コンピューター システムの命令のシーケンスを解釈できるデバイスを表します。
'* マルチプロセッサ コンピューターでは、各プロセッサにこのクラスのインスタンスが
'*  1 つあります。
'* WMI Provider : CIMWin32
'* UUID : {8502C4BB-5FBB-11D2-AAC1-006008C78BC7}
'*
'* @see http://www.wmifun.net/library/win32_processor.html
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32Processor = "Win32_Processor"

'*-----------------------------------------------------------------------------
'* Win32_LogicalDisk クラス
'* Win32 システム上の実際のローカル記憶域デバイスに解決するデータ ソースを表し
'* ます。クラスはローカルとマップされた論理ディスクの両方を返します。
'* ただし、ローカル ディスクの情報取得にはこのクラス、マップされた論理ディスクの
'* 情報取得には Win32_MappedLogicalDisk を使用する対処法が推奨されています。
'* WMI Provider : CIMWin32
'* UUID : {8502C4B7-5FBB-11D2-AAC1-006008C78BC7}
'*
'* @see http://www.wmifun.net/library/win32_logicaldisk.html
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32LogicalDisk = "Win32_LogicalDisk"

'*-----------------------------------------------------------------------------
'* Win32_NetworkAdapter クラス
'* Win32 システム上のネットワーク アダプターを表します。
'* WMI Provider : CIMWin32
'* UUID : {8502C4C0-5FBB-11D2-AAC1-006008C78BC7}
'*
'* @see http://www.wmifun.net/library/win32_networkadapter.html
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32NetworkAdapter = "Win32_NetworkAdapter"

'*-----------------------------------------------------------------------------
'* Win32_PnPEntity クラス
'* プラグ アンド プレイ デバイスのプロパティを表します。
'* プラグ アンド プレイのエンティティはコントロール パネルにあるデバイス
'* マネージャーのエンティティとして表示されます。
'* WMI Provider : CIMWin32
'* UUID : {FE28FD98-C875-11d2-B352-00104BC97924}
'*
'* @see http://www.wmifun.net/library/win32_pnpentity.html
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PnPEntity = "Win32_PnPEntity"

'*-----------------------------------------------------------------------------
'* Win32_Service クラス
'* Win32 コンピューター システムのサービスを表します。
'* サービス アプリケーションは、サービス コントロール マネージャー (SCM) の
'* インターフェイス規則に適合し、サービス コントロール パネル ユーティリティで
'* システム起動時に自動的にユーザーによってか、または Win32 API に含まれるサー
'* ビス機能を使用するアプリケーションによって開始されます。
'* システムにログオンしているユーザーがいないときでもサービスを実行できます。
'* WMI Provider : CIMWin32
'* UUID : {8502C4D9-5FBB-11D2-AAC1-006008C78BC7}
'*
'* @see http://www.wmifun.net/library/win32_service.html
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32Service = "Win32_Service"

'*-----------------------------------------------------------------------------
'* Win32_DiskDrive クラス
'* Win32 オペレーティング システムを実行しているコンピューターで見ることができる
'* 物理ディスク ドライブを表します。
'* Win32 物理ディスク ドライブへのインターフェイスは、このクラスの下層 (または
'* メンバー) です。このオブジェクトから見ることができるディスク ドライブの機能
'* は、ドライブの論理特性および管理特性に対応しています。
'* デバイスの実際の物理的特性に影響しない場合もあります。
'* 別の論理デバイスに基づいたオブジェクトは、このクラスのメンバーではありません。
'* WMI Provider : CIMWin32
'* UUID : {8502C4B2-5FBB-11D2-AAC1-006008C78BC7}
'*
'* @see http://www.wmifun.net/library/win32_diskdrive.html
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32DiskDrive = "Win32_DiskDrive"

'*-----------------------------------------------------------------------------
'* Win32_PingStatus クラス
'* 標準 ping コマンドによって返される値が含まれます。
'* ping の詳細は RFC 791 にあります。
'* WMI Provider : WMIPingProvider
'* UUID :
'*
'* @see http://www.wmifun.net/library/win32_pingstatus.html
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PingStatus = "Win32_PingStatus"

'*-----------------------------------------------------------------------------
'* Win32_UserAccount クラス
'* Win32 システムのユーザー アカウントに関する情報が含まれています。
'* WMI Provider : CIMWin32
'* UUID : {8502C4CC-5FBB-11D2-AAC1-006008C78BC7}
'*
'* @see http://www.wmifun.net/library/win32_useraccount.html
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32UserAccount = "Win32_UserAccount"

'*-----------------------------------------------------------------------------
'* StdRegProv クラス
'* システム レジストリと対話するメソッドを含みます。
'* これらのメソッドを使用して次のことを実行できます:
'*   ・ユーザーのアクセス許可の確認
'*   ・レジストリの作成、列挙、および削除の実行
'*   ・名前を付けた値の作成、列挙、および削除
'*   ・データ値の読み取り、書き込み、および削除
'* WMI Provider : RegProv
'* UUID :
'*
'* @see http://www.wmifun.net/library/stdregprov.html
'*-----------------------------------------------------------------------------
Const WmiClassNameStdRegProv = "StdRegProv"

'*-----------------------------------------------------------------------------
'* CIM_DataFile クラス
'*
'*
'* WMI Provider : CIMWin32
'* UUID : {8502C55A-5FBB-11D2-AAC1-006008C78BC7}
'*
'* @see http://www.wmifun.net/library/cim_datafile.html
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMDataFile = "CIM_DataFile"

'*-----------------------------------------------------------------------------
'* Win32_Product クラス
'* MSI でインストールされる製品を表します。製品は通常、単一のインストール
'* パッケージと相関しています。
'* WMI Provider : MSIProv
'* UUID : {CE3324AA-DB34-11d2-85FC-0000F8102E5F}
'*
'* @see http://www.wmifun.net/library/win32_product.html
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32Product = "Win32_Product"

'*-----------------------------------------------------------------------------
'* Win32_ScheduledJob クラス
'* ネットワーク管理スケジュール サービス機能 ("ジョブ" および "AT コマンド"
'* 機能) を使用するスケジュールされたジョブを表します。
'* これは、Windows 2000 タスク スケジューラを使ってスケジュールしたタスクと
'* 異なることに注意してください。
'* このクラスは Windows NT 4.0 以降でのみ使用されます。
'* スケジュール サービスに対してスケジュールされた各ジョブは、持続的に格納され
'*  (スケジューラにより再起動後もジョブの開始が認知されます)、週と月の指定され
'* た日時に実行されます。
'* コンピューターが動作していないか、またはスケジュール サービスが指定された
'* ジョブ時間に実行していない場合、スケジュール サービスにより指定された時間で
'* 次の日に指定されたジョブが実行されます。
'* スケジュールされたジョブは、協定世界時 (UTC) に関連して、たとえば GMT からの
'* ずれオフセットでスケジュールされます。これは、タイム ゾーン仕様を使用して
'* ジョブを指定できることを意味します。Win32_ScheduledJob により、オブジェクト
'* が列挙されるときに UTC オフセットでローカル タイムが返され、新しいジョブが
'* 作成されるときにローカル タイムに変換されます。
'* たとえば、ボストンで太平洋標準時で月曜の午後 10:30 に実行するように指定され
'* たジョブは、ローカルでは東部標準時で火曜の午前 1:30 に実行されるようにスケ
'* ジュールされます。
'* 夏時間でローカル コンピューターを操作しているかどうかをクライアントは考慮す
'* る必要があることに注意し、操作している場合 UTC オフセットから 60 分のずれを
'* 引いてください。
'* WMI Provider : CIMWin32
'* UUID : {8502C4E0-5FBB-11D2-AAC1-006008C78BC7}
'*
'* @see http://www.wmifun.net/library/win32_scheduledjob.html
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ScheduledJob = "Win32_ScheduledJob"

'*-----------------------------------------------------------------------------
'* Win32_ComputerSystemProduct クラス
'* 製品を表します。これは、このコンピューター システムで使用されるソフトウェア
'* とハードウェアが含まれます。
'* WMI Provider : CIMWin32
'* UUID : {FAF76B96-798C-11D2-AAD1-006008C78BC7}
'*
'* @see http://www.wmifun.net/library/win32_computersystemproduct.html
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ComputerSystemProduct = "Win32_ComputerSystemProduct"

'*-----------------------------------------------------------------------------
'* Win32_BIOS クラス
'* コンピューターにインストールされている基本入出力 (BIOS) の属性を表します。
'* WMI Provider : CIMWin32
'* UUID : {8502C4E1-5FBB-11D2-AAC1-006008C78BC7}
'*
'* @see http://www.wmifun.net/library/win32_bios.html
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32BIOS = "Win32_BIOS"
'*-----------------------------------------------------------------------------
'* CIM_Action クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMAction = "CIM_Action"

'*-----------------------------------------------------------------------------
'* CIM_ActionSequence クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMActionSequence = "CIM_ActionSequence"

'*-----------------------------------------------------------------------------
'* CIM_ActsAsSpare クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMActsAsSpare = "CIM_ActsAsSpare"

'*-----------------------------------------------------------------------------
'* CIM_AdjacentSlots クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMAdjacentSlots = "CIM_AdjacentSlots"

'*-----------------------------------------------------------------------------
'* CIM_AggregatePExtent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMAggregatePExtent = "CIM_AggregatePExtent"

'*-----------------------------------------------------------------------------
'* CIM_AggregatePSExtent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMAggregatePSExtent = "CIM_AggregatePSExtent"

'*-----------------------------------------------------------------------------
'* CIM_AggregateRedundancyComponent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMAggregateRedundancyComponent = "CIM_AggregateRedundancyComponent"

'*-----------------------------------------------------------------------------
'* CIM_AlarmDevice クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMAlarmDevice = "CIM_AlarmDevice"

'*-----------------------------------------------------------------------------
'* CIM_AllocatedResource クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMAllocatedResource = "CIM_AllocatedResource"

'*-----------------------------------------------------------------------------
'* CIM_ApplicationSystem クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMApplicationSystem = "CIM_ApplicationSystem"

'*-----------------------------------------------------------------------------
'* CIM_ApplicationSystemSoftwareFeature クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMApplicationSystemSoftwareFeature = "CIM_ApplicationSystemSoftwareFeature"

'*-----------------------------------------------------------------------------
'* CIM_AssociatedAlarm クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMAssociatedAlarm = "CIM_AssociatedAlarm"

'*-----------------------------------------------------------------------------
'* CIM_AssociatedBattery クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMAssociatedBattery = "CIM_AssociatedBattery"

'*-----------------------------------------------------------------------------
'* CIM_AssociatedCooling クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMAssociatedCooling = "CIM_AssociatedCooling"

'*-----------------------------------------------------------------------------
'* CIM_AssociatedMemory クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMAssociatedMemory = "CIM_AssociatedMemory"

'*-----------------------------------------------------------------------------
'* CIM_AssociatedProcessorMemory クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMAssociatedProcessorMemory = "CIM_AssociatedProcessorMemory"

'*-----------------------------------------------------------------------------
'* CIM_AssociatedSensor クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMAssociatedSensor = "CIM_AssociatedSensor"

'*-----------------------------------------------------------------------------
'* CIM_AssociatedSupplyCurrentSensor クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMAssociatedSupplyCurrentSensor = "CIM_AssociatedSupplyCurrentSensor"

'*-----------------------------------------------------------------------------
'* CIM_AssociatedSupplyVoltageSensor クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMAssociatedSupplyVoltageSensor = "CIM_AssociatedSupplyVoltageSensor"

'*-----------------------------------------------------------------------------
'* CIM_BasedOn クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMBasedOn = "CIM_BasedOn"

'*-----------------------------------------------------------------------------
'* CIM_Battery クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMBattery = "CIM_Battery"

'*-----------------------------------------------------------------------------
'* CIM_BinarySensor クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMBinarySensor = "CIM_BinarySensor"

'*-----------------------------------------------------------------------------
'* CIM_BIOSElement クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMBIOSElement = "CIM_BIOSElement"

'*-----------------------------------------------------------------------------
'* CIM_BIOSFeature クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMBIOSFeature = "CIM_BIOSFeature"

'*-----------------------------------------------------------------------------
'* CIM_BIOSFeatureBIOSElements クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMBIOSFeatureBIOSElements = "CIM_BIOSFeatureBIOSElements"

'*-----------------------------------------------------------------------------
'* CIM_BIOSLoadedInNV クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMBIOSLoadedInNV = "CIM_BIOSLoadedInNV"

'*-----------------------------------------------------------------------------
'* CIM_BootOSFromFS クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMBootOSFromFS = "CIM_BootOSFromFS"

'*-----------------------------------------------------------------------------
'* CIM_BootSAP クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMBootSAP = "CIM_BootSAP"

'*-----------------------------------------------------------------------------
'* CIM_BootService クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMBootService = "CIM_BootService"

'*-----------------------------------------------------------------------------
'* CIM_BootServiceAccessBySAP クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMBootServiceAccessBySAP = "CIM_BootServiceAccessBySAP"

'*-----------------------------------------------------------------------------
'* CIM_CacheMemory クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMCacheMemory = "CIM_CacheMemory"

'*-----------------------------------------------------------------------------
'* CIM_Card クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMCard = "CIM_Card"

'*-----------------------------------------------------------------------------
'* CIM_CardInSlot クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMCardInSlot = "CIM_CardInSlot"

'*-----------------------------------------------------------------------------
'* CIM_CardOnCard クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMCardOnCard = "CIM_CardOnCard"

'*-----------------------------------------------------------------------------
'* CIM_CDROMDrive クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMCDROMDrive = "CIM_CDROMDrive"

'*-----------------------------------------------------------------------------
'* CIM_Chassis クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMChassis = "CIM_Chassis"

'*-----------------------------------------------------------------------------
'* CIM_ChassisInRack クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMChassisInRack = "CIM_ChassisInRack"

'*-----------------------------------------------------------------------------
'* CIM_Check クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMCheck = "CIM_Check"

'*-----------------------------------------------------------------------------
'* CIM_Chip クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMChip = "CIM_Chip"

'*-----------------------------------------------------------------------------
'* CIM_ClusteringSAP クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMClusteringSAP = "CIM_ClusteringSAP"

'*-----------------------------------------------------------------------------
'* CIM_ClusteringService クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMClusteringService = "CIM_ClusteringService"

'*-----------------------------------------------------------------------------
'* CIM_ClusterServiceAccessBySAP クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMClusterServiceAccessBySAP = "CIM_ClusterServiceAccessBySAP"

'*-----------------------------------------------------------------------------
'* CIM_CollectedCollections クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMCollectedCollections = "CIM_CollectedCollections"

'*-----------------------------------------------------------------------------
'* CIM_CollectedMSEs クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMCollectedMSEs = "CIM_CollectedMSEs"

'*-----------------------------------------------------------------------------
'* CIM_CollectionOfMSEs クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMCollectionOfMSEs = "CIM_CollectionOfMSEs"

'*-----------------------------------------------------------------------------
'* CIM_CollectionOfSensors クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMCollectionOfSensors = "CIM_CollectionOfSensors"

'*-----------------------------------------------------------------------------
'* CIM_CollectionSetting クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMCollectionSetting = "CIM_CollectionSetting"

'*-----------------------------------------------------------------------------
'* CIM_CompatibleProduct クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMCompatibleProduct = "CIM_CompatibleProduct"

'*-----------------------------------------------------------------------------
'* CIM_Component クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMComponent = "CIM_Component"

'*-----------------------------------------------------------------------------
'* CIM_ComputerSystem クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMComputerSystem = "CIM_ComputerSystem"

'*-----------------------------------------------------------------------------
'* CIM_ComputerSystemDMA クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMComputerSystemDMA = "CIM_ComputerSystemDMA"

'*-----------------------------------------------------------------------------
'* CIM_ComputerSystemIRQ クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMComputerSystemIRQ = "CIM_ComputerSystemIRQ"

'*-----------------------------------------------------------------------------
'* CIM_ComputerSystemMappedIO クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMComputerSystemMappedIO = "CIM_ComputerSystemMappedIO"

'*-----------------------------------------------------------------------------
'* CIM_ComputerSystemPackage クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMComputerSystemPackage = "CIM_ComputerSystemPackage"

'*-----------------------------------------------------------------------------
'* CIM_ComputerSystemResource クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMComputerSystemResource = "CIM_ComputerSystemResource"

'*-----------------------------------------------------------------------------
'* CIM_Configuration クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMConfiguration = "CIM_Configuration"

'*-----------------------------------------------------------------------------
'* CIM_ConnectedTo クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMConnectedTo = "CIM_ConnectedTo"

'*-----------------------------------------------------------------------------
'* CIM_ConnectorOnPackage クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMConnectorOnPackage = "CIM_ConnectorOnPackage"

'*-----------------------------------------------------------------------------
'* CIM_Container クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMContainer = "CIM_Container"

'*-----------------------------------------------------------------------------
'* CIM_ControlledBy クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMControlledBy = "CIM_ControlledBy"

'*-----------------------------------------------------------------------------
'* CIM_Controller クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMController = "CIM_Controller"

'*-----------------------------------------------------------------------------
'* CIM_CoolingDevice クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMCoolingDevice = "CIM_CoolingDevice"

'*-----------------------------------------------------------------------------
'* CIM_CopyFileAction クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMCopyFileAction = "CIM_CopyFileAction"

'*-----------------------------------------------------------------------------
'* CIM_CreateDirectoryAction クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMCreateDirectoryAction = "CIM_CreateDirectoryAction"

'*-----------------------------------------------------------------------------
'* CIM_CurrentSensor クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMCurrentSensor = "CIM_CurrentSensor"

'*-----------------------------------------------------------------------------
'* CIM_Dependency クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMDependency = "CIM_Dependency"

'*-----------------------------------------------------------------------------
'* CIM_DependencyContext クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMDependencyContext = "CIM_DependencyContext"

'*-----------------------------------------------------------------------------
'* CIM_DesktopMonitor クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMDesktopMonitor = "CIM_DesktopMonitor"

'*-----------------------------------------------------------------------------
'* CIM_DeviceAccessedByFile クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMDeviceAccessedByFile = "CIM_DeviceAccessedByFile"

'*-----------------------------------------------------------------------------
'* CIM_DeviceConnection クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMDeviceConnection = "CIM_DeviceConnection"

'*-----------------------------------------------------------------------------
'* CIM_DeviceErrorCounts クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMDeviceErrorCounts = "CIM_DeviceErrorCounts"

'*-----------------------------------------------------------------------------
'* CIM_DeviceFile クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMDeviceFile = "CIM_DeviceFile"

'*-----------------------------------------------------------------------------
'* CIM_DeviceSAPImplementation クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMDeviceSAPImplementation = "CIM_DeviceSAPImplementation"

'*-----------------------------------------------------------------------------
'* CIM_DeviceServiceImplementation クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMDeviceServiceImplementation = "CIM_DeviceServiceImplementation"

'*-----------------------------------------------------------------------------
'* CIM_DeviceSoftware クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMDeviceSoftware = "CIM_DeviceSoftware"

'*-----------------------------------------------------------------------------
'* CIM_Directory クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMDirectory = "CIM_Directory"

'*-----------------------------------------------------------------------------
'* CIM_DirectoryAction クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMDirectoryAction = "CIM_DirectoryAction"

'*-----------------------------------------------------------------------------
'* CIM_DirectoryContainsFile クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMDirectoryContainsFile = "CIM_DirectoryContainsFile"

'*-----------------------------------------------------------------------------
'* CIM_DirectorySpecification クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMDirectorySpecification = "CIM_DirectorySpecification"

'*-----------------------------------------------------------------------------
'* CIM_DirectorySpecificationFile クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMDirectorySpecificationFile = "CIM_DirectorySpecificationFile"

'*-----------------------------------------------------------------------------
'* CIM_DiscreteSensor クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMDiscreteSensor = "CIM_DiscreteSensor"

'*-----------------------------------------------------------------------------
'* CIM_DiskDrive クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMDiskDrive = "CIM_DiskDrive"

'*-----------------------------------------------------------------------------
'* CIM_DisketteDrive クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMDisketteDrive = "CIM_DisketteDrive"

'*-----------------------------------------------------------------------------
'* CIM_DiskPartition クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMDiskPartition = "CIM_DiskPartition"

'*-----------------------------------------------------------------------------
'* CIM_DiskSpaceCheck クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMDiskSpaceCheck = "CIM_DiskSpaceCheck"

'*-----------------------------------------------------------------------------
'* CIM_Display クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMDisplay = "CIM_Display"

'*-----------------------------------------------------------------------------
'* CIM_DMA クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMDMA = "CIM_DMA"

'*-----------------------------------------------------------------------------
'* CIM_Docked クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMDocked = "CIM_Docked"

'*-----------------------------------------------------------------------------
'* CIM_ElementCapacity クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMElementCapacity = "CIM_ElementCapacity"

'*-----------------------------------------------------------------------------
'* CIM_ElementConfiguration クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMElementConfiguration = "CIM_ElementConfiguration"

'*-----------------------------------------------------------------------------
'* CIM_ElementSetting クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMElementSetting = "CIM_ElementSetting"

'*-----------------------------------------------------------------------------
'* CIM_ElementsLinked クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMElementsLinked = "CIM_ElementsLinked"

'*-----------------------------------------------------------------------------
'* CIM_ErrorCountersForDevice クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMErrorCountersForDevice = "CIM_ErrorCountersForDevice"

'*-----------------------------------------------------------------------------
'* CIM_ExecuteProgram クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMExecuteProgram = "CIM_ExecuteProgram"

'*-----------------------------------------------------------------------------
'* CIM_Export クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMExport = "CIM_Export"

'*-----------------------------------------------------------------------------
'* CIM_ExtraCapacityGroup クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMExtraCapacityGroup = "CIM_ExtraCapacityGroup"

'*-----------------------------------------------------------------------------
'* CIM_Fan クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMFan = "CIM_Fan"

'*-----------------------------------------------------------------------------
'* CIM_FileAction クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMFileAction = "CIM_FileAction"

'*-----------------------------------------------------------------------------
'* CIM_FileSpecification クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMFileSpecification = "CIM_FileSpecification"

'*-----------------------------------------------------------------------------
'* CIM_FileStorage クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMFileStorage = "CIM_FileStorage"

'*-----------------------------------------------------------------------------
'* CIM_FileSystem クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMFileSystem = "CIM_FileSystem"

'*-----------------------------------------------------------------------------
'* CIM_FlatPanel クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMFlatPanel = "CIM_FlatPanel"

'*-----------------------------------------------------------------------------
'* CIM_FromDirectoryAction クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMFromDirectoryAction = "CIM_FromDirectoryAction"

'*-----------------------------------------------------------------------------
'* CIM_FromDirectorySpecification クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMFromDirectorySpecification = "CIM_FromDirectorySpecification"

'*-----------------------------------------------------------------------------
'* CIM_FRU クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMFRU = "CIM_FRU"

'*-----------------------------------------------------------------------------
'* CIM_FRUIncludesProduct クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMFRUIncludesProduct = "CIM_FRUIncludesProduct"

'*-----------------------------------------------------------------------------
'* CIM_FRUPhysicalElements クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMFRUPhysicalElements = "CIM_FRUPhysicalElements"

'*-----------------------------------------------------------------------------
'* CIM_HeatPipe クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMHeatPipe = "CIM_HeatPipe"

'*-----------------------------------------------------------------------------
'* CIM_HostedAccessPoint クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMHostedAccessPoint = "CIM_HostedAccessPoint"

'*-----------------------------------------------------------------------------
'* CIM_HostedBootSAP クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMHostedBootSAP = "CIM_HostedBootSAP"

'*-----------------------------------------------------------------------------
'* CIM_HostedBootService クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMHostedBootService = "CIM_HostedBootService"

'*-----------------------------------------------------------------------------
'* CIM_HostedFileSystem クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMHostedFileSystem = "CIM_HostedFileSystem"

'*-----------------------------------------------------------------------------
'* CIM_HostedJobDestination クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMHostedJobDestination = "CIM_HostedJobDestination"

'*-----------------------------------------------------------------------------
'* CIM_HostedService クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMHostedService = "CIM_HostedService"

'*-----------------------------------------------------------------------------
'* CIM_InfraredController クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMInfraredController = "CIM_InfraredController"

'*-----------------------------------------------------------------------------
'* CIM_InstalledOS クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMInstalledOS = "CIM_InstalledOS"

'*-----------------------------------------------------------------------------
'* CIM_InstalledSoftwareElement クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMInstalledSoftwareElement = "CIM_InstalledSoftwareElement"

'*-----------------------------------------------------------------------------
'* CIM_IRQ クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMIRQ = "CIM_IRQ"

'*-----------------------------------------------------------------------------
'* CIM_Job クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMJob = "CIM_Job"

'*-----------------------------------------------------------------------------
'* CIM_JobDestination クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMJobDestination = "CIM_JobDestination"

'*-----------------------------------------------------------------------------
'* CIM_JobDestinationJobs クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMJobDestinationJobs = "CIM_JobDestinationJobs"

'*-----------------------------------------------------------------------------
'* CIM_Keyboard クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMKeyboard = "CIM_Keyboard"

'*-----------------------------------------------------------------------------
'* CIM_LinkHasConnector クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMLinkHasConnector = "CIM_LinkHasConnector"

'*-----------------------------------------------------------------------------
'* CIM_LocalFileSystem クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMLocalFileSystem = "CIM_LocalFileSystem"

'*-----------------------------------------------------------------------------
'* CIM_Location クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMLocation = "CIM_Location"

'*-----------------------------------------------------------------------------
'* CIM_LogicalDevice クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMLogicalDevice = "CIM_LogicalDevice"

'*-----------------------------------------------------------------------------
'* CIM_LogicalDisk クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMLogicalDisk = "CIM_LogicalDisk"

'*-----------------------------------------------------------------------------
'* CIM_LogicalDiskBasedOnPartition クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMLogicalDiskBasedOnPartition = "CIM_LogicalDiskBasedOnPartition"

'*-----------------------------------------------------------------------------
'* CIM_LogicalDiskBasedOnVolumeSet クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMLogicalDiskBasedOnVolumeSet = "CIM_LogicalDiskBasedOnVolumeSet"

'*-----------------------------------------------------------------------------
'* CIM_LogicalElement クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMLogicalElement = "CIM_LogicalElement"

'*-----------------------------------------------------------------------------
'* CIM_LogicalFile クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMLogicalFile = "CIM_LogicalFile"

'*-----------------------------------------------------------------------------
'* CIM_LogicalIdentity クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMLogicalIdentity = "CIM_LogicalIdentity"

'*-----------------------------------------------------------------------------
'* CIM_MagnetoOpticalDrive クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMMagnetoOpticalDrive = "CIM_MagnetoOpticalDrive"

'*-----------------------------------------------------------------------------
'* CIM_ManagedSystemElement クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMManagedSystemElement = "CIM_ManagedSystemElement"

'*-----------------------------------------------------------------------------
'* CIM_ManagementController クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMManagementController = "CIM_ManagementController"

'*-----------------------------------------------------------------------------
'* CIM_MediaAccessDevice クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMMediaAccessDevice = "CIM_MediaAccessDevice"

'*-----------------------------------------------------------------------------
'* CIM_MediaPresent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMMediaPresent = "CIM_MediaPresent"

'*-----------------------------------------------------------------------------
'* CIM_Memory クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMMemory = "CIM_Memory"

'*-----------------------------------------------------------------------------
'* CIM_MemoryCapacity クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMMemoryCapacity = "CIM_MemoryCapacity"

'*-----------------------------------------------------------------------------
'* CIM_MemoryCheck クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMMemoryCheck = "CIM_MemoryCheck"

'*-----------------------------------------------------------------------------
'* CIM_MemoryMappedIO クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMMemoryMappedIO = "CIM_MemoryMappedIO"

'*-----------------------------------------------------------------------------
'* CIM_MemoryOnCard クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMMemoryOnCard = "CIM_MemoryOnCard"

'*-----------------------------------------------------------------------------
'* CIM_MemoryWithMedia クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMMemoryWithMedia = "CIM_MemoryWithMedia"

'*-----------------------------------------------------------------------------
'* CIM_ModifySettingAction クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMModifySettingAction = "CIM_ModifySettingAction"

'*-----------------------------------------------------------------------------
'* CIM_MonitorResolution クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMMonitorResolution = "CIM_MonitorResolution"

'*-----------------------------------------------------------------------------
'* CIM_MonitorSetting クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMMonitorSetting = "CIM_MonitorSetting"

'*-----------------------------------------------------------------------------
'* CIM_Mount クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMMount = "CIM_Mount"

'*-----------------------------------------------------------------------------
'* CIM_MultiStateSensor クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMMultiStateSensor = "CIM_MultiStateSensor"

'*-----------------------------------------------------------------------------
'* CIM_NetworkAdapter クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMNetworkAdapter = "CIM_NetworkAdapter"

'*-----------------------------------------------------------------------------
'* CIM_NFS クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMNFS = "CIM_NFS"

'*-----------------------------------------------------------------------------
'* CIM_NonVolatileStorage クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMNonVolatileStorage = "CIM_NonVolatileStorage"

'*-----------------------------------------------------------------------------
'* CIM_NumericSensor クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMNumericSensor = "CIM_NumericSensor"

'*-----------------------------------------------------------------------------
'* CIM_OperatingSystem クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMOperatingSystem = "CIM_OperatingSystem"

'*-----------------------------------------------------------------------------
'* CIM_OperatingSystemSoftwareFeature クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMOperatingSystemSoftwareFeature = "CIM_OperatingSystemSoftwareFeature"

'*-----------------------------------------------------------------------------
'* CIM_OSProcess クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMOSProcess = "CIM_OSProcess"

'*-----------------------------------------------------------------------------
'* CIM_OSVersionCheck クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMOSVersionCheck = "CIM_OSVersionCheck"

'*-----------------------------------------------------------------------------
'* CIM_PackageAlarm クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMPackageAlarm = "CIM_PackageAlarm"

'*-----------------------------------------------------------------------------
'* CIM_PackageCooling クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMPackageCooling = "CIM_PackageCooling"

'*-----------------------------------------------------------------------------
'* CIM_PackagedComponent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMPackagedComponent = "CIM_PackagedComponent"

'*-----------------------------------------------------------------------------
'* CIM_PackageInChassis クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMPackageInChassis = "CIM_PackageInChassis"

'*-----------------------------------------------------------------------------
'* CIM_PackageInSlot クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMPackageInSlot = "CIM_PackageInSlot"

'*-----------------------------------------------------------------------------
'* CIM_PackageTempSensor クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMPackageTempSensor = "CIM_PackageTempSensor"

'*-----------------------------------------------------------------------------
'* CIM_ParallelController クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMParallelController = "CIM_ParallelController"

'*-----------------------------------------------------------------------------
'* CIM_ParticipatesInSet クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMParticipatesInSet = "CIM_ParticipatesInSet"

'*-----------------------------------------------------------------------------
'* CIM_PCIController クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMPCIController = "CIM_PCIController"

'*-----------------------------------------------------------------------------
'* CIM_PCMCIAController クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMPCMCIAController = "CIM_PCMCIAController"

'*-----------------------------------------------------------------------------
'* CIM_PCVideoController クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMPCVideoController = "CIM_PCVideoController"

'*-----------------------------------------------------------------------------
'* CIM_PExtentRedundancyComponent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMPExtentRedundancyComponent = "CIM_PExtentRedundancyComponent"

'*-----------------------------------------------------------------------------
'* CIM_PhysicalCapacity クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMPhysicalCapacity = "CIM_PhysicalCapacity"

'*-----------------------------------------------------------------------------
'* CIM_PhysicalComponent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMPhysicalComponent = "CIM_PhysicalComponent"

'*-----------------------------------------------------------------------------
'* CIM_PhysicalConnector クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMPhysicalConnector = "CIM_PhysicalConnector"

'*-----------------------------------------------------------------------------
'* CIM_PhysicalElement クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMPhysicalElement = "CIM_PhysicalElement"

'*-----------------------------------------------------------------------------
'* CIM_PhysicalElementLocation クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMPhysicalElementLocation = "CIM_PhysicalElementLocation"

'*-----------------------------------------------------------------------------
'* CIM_PhysicalExtent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMPhysicalExtent = "CIM_PhysicalExtent"

'*-----------------------------------------------------------------------------
'* CIM_PhysicalFrame クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMPhysicalFrame = "CIM_PhysicalFrame"

'*-----------------------------------------------------------------------------
'* CIM_PhysicalLink クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMPhysicalLink = "CIM_PhysicalLink"

'*-----------------------------------------------------------------------------
'* CIM_PhysicalMedia クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMPhysicalMedia = "CIM_PhysicalMedia"

'*-----------------------------------------------------------------------------
'* CIM_PhysicalMemory クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMPhysicalMemory = "CIM_PhysicalMemory"

'*-----------------------------------------------------------------------------
'* CIM_PhysicalPackage クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMPhysicalPackage = "CIM_PhysicalPackage"

'*-----------------------------------------------------------------------------
'* CIM_PointingDevice クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMPointingDevice = "CIM_PointingDevice"

'*-----------------------------------------------------------------------------
'* CIM_PotsModem クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMPotsModem = "CIM_PotsModem"

'*-----------------------------------------------------------------------------
'* CIM_PowerSupply クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMPowerSupply = "CIM_PowerSupply"

'*-----------------------------------------------------------------------------
'* CIM_Printer クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMPrinter = "CIM_Printer"

'*-----------------------------------------------------------------------------
'* CIM_Process クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMProcess = "CIM_Process"

'*-----------------------------------------------------------------------------
'* CIM_ProcessExecutable クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMProcessExecutable = "CIM_ProcessExecutable"

'*-----------------------------------------------------------------------------
'* CIM_Processor クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMProcessor = "CIM_Processor"

'*-----------------------------------------------------------------------------
'* CIM_ProcessThread クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMProcessThread = "CIM_ProcessThread"

'*-----------------------------------------------------------------------------
'* CIM_Product クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMProduct = "CIM_Product"

'*-----------------------------------------------------------------------------
'* CIM_ProductFRU クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMProductFRU = "CIM_ProductFRU"

'*-----------------------------------------------------------------------------
'* CIM_ProductParentChild クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMProductParentChild = "CIM_ProductParentChild"

'*-----------------------------------------------------------------------------
'* CIM_ProductPhysicalElements クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMProductPhysicalElements = "CIM_ProductPhysicalElements"

'*-----------------------------------------------------------------------------
'* CIM_ProductProductDependency クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMProductProductDependency = "CIM_ProductProductDependency"

'*-----------------------------------------------------------------------------
'* CIM_ProductSoftwareFeatures クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMProductSoftwareFeatures = "CIM_ProductSoftwareFeatures"

'*-----------------------------------------------------------------------------
'* CIM_ProductSupport クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMProductSupport = "CIM_ProductSupport"

'*-----------------------------------------------------------------------------
'* CIM_ProtectedSpaceExtent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMProtectedSpaceExtent = "CIM_ProtectedSpaceExtent"

'*-----------------------------------------------------------------------------
'* CIM_PSExtentBasedOnPExtent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMPSExtentBasedOnPExtent = "CIM_PSExtentBasedOnPExtent"

'*-----------------------------------------------------------------------------
'* CIM_Rack クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMRack = "CIM_Rack"

'*-----------------------------------------------------------------------------
'* CIM_Realizes クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMRealizes = "CIM_Realizes"

'*-----------------------------------------------------------------------------
'* CIM_RealizesAggregatePExtent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMRealizesAggregatePExtent = "CIM_RealizesAggregatePExtent"

'*-----------------------------------------------------------------------------
'* CIM_RealizesDiskPartition クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMRealizesDiskPartition = "CIM_RealizesDiskPartition"

'*-----------------------------------------------------------------------------
'* CIM_RealizesPExtent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMRealizesPExtent = "CIM_RealizesPExtent"

'*-----------------------------------------------------------------------------
'* CIM_RebootAction クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMRebootAction = "CIM_RebootAction"

'*-----------------------------------------------------------------------------
'* CIM_RedundancyComponent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMRedundancyComponent = "CIM_RedundancyComponent"

'*-----------------------------------------------------------------------------
'* CIM_RedundancyGroup クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMRedundancyGroup = "CIM_RedundancyGroup"

'*-----------------------------------------------------------------------------
'* CIM_Refrigeration クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMRefrigeration = "CIM_Refrigeration"

'*-----------------------------------------------------------------------------
'* CIM_RelatedStatistics クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMRelatedStatistics = "CIM_RelatedStatistics"

'*-----------------------------------------------------------------------------
'* CIM_RemoteFileSystem クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMRemoteFileSystem = "CIM_RemoteFileSystem"

'*-----------------------------------------------------------------------------
'* CIM_RemoveDirectoryAction クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMRemoveDirectoryAction = "CIM_RemoveDirectoryAction"

'*-----------------------------------------------------------------------------
'* CIM_RemoveFileAction クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMRemoveFileAction = "CIM_RemoveFileAction"

'*-----------------------------------------------------------------------------
'* CIM_ReplacementSet クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMReplacementSet = "CIM_ReplacementSet"

'*-----------------------------------------------------------------------------
'* CIM_ResidesOnExtent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMResidesOnExtent = "CIM_ResidesOnExtent"

'*-----------------------------------------------------------------------------
'* CIM_RunningOS クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMRunningOS = "CIM_RunningOS"

'*-----------------------------------------------------------------------------
'* CIM_SAPSAPDependency クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMSAPSAPDependency = "CIM_SAPSAPDependency"

'*-----------------------------------------------------------------------------
'* CIM_Scanner クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMScanner = "CIM_Scanner"

'*-----------------------------------------------------------------------------
'* CIM_SCSIController クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMSCSIController = "CIM_SCSIController"

'*-----------------------------------------------------------------------------
'* CIM_SCSIInterface クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMSCSIInterface = "CIM_SCSIInterface"

'*-----------------------------------------------------------------------------
'* CIM_Sensor クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMSensor = "CIM_Sensor"

'*-----------------------------------------------------------------------------
'* CIM_SerialController クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMSerialController = "CIM_SerialController"

'*-----------------------------------------------------------------------------
'* CIM_SerialInterface クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMSerialInterface = "CIM_SerialInterface"

'*-----------------------------------------------------------------------------
'* CIM_Service クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMService = "CIM_Service"

'*-----------------------------------------------------------------------------
'* CIM_ServiceAccessBySAP クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMServiceAccessBySAP = "CIM_ServiceAccessBySAP"

'*-----------------------------------------------------------------------------
'* CIM_ServiceAccessPoint クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMServiceAccessPoint = "CIM_ServiceAccessPoint"

'*-----------------------------------------------------------------------------
'* CIM_ServiceSAPDependency クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMServiceSAPDependency = "CIM_ServiceSAPDependency"

'*-----------------------------------------------------------------------------
'* CIM_ServiceServiceDependency クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMServiceServiceDependency = "CIM_ServiceServiceDependency"

'*-----------------------------------------------------------------------------
'* CIM_Setting クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMSetting = "CIM_Setting"

'*-----------------------------------------------------------------------------
'* CIM_SettingCheck クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMSettingCheck = "CIM_SettingCheck"

'*-----------------------------------------------------------------------------
'* CIM_SettingContext クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMSettingContext = "CIM_SettingContext"

'*-----------------------------------------------------------------------------
'* CIM_Slot クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMSlot = "CIM_Slot"

'*-----------------------------------------------------------------------------
'* CIM_SlotInSlot クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMSlotInSlot = "CIM_SlotInSlot"

'*-----------------------------------------------------------------------------
'* CIM_SoftwareElement クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMSoftwareElement = "CIM_SoftwareElement"

'*-----------------------------------------------------------------------------
'* CIM_SoftwareElementActions クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMSoftwareElementActions = "CIM_SoftwareElementActions"

'*-----------------------------------------------------------------------------
'* CIM_SoftwareElementChecks クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMSoftwareElementChecks = "CIM_SoftwareElementChecks"

'*-----------------------------------------------------------------------------
'* CIM_SoftwareElementVersionCheck クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMSoftwareElementVersionCheck = "CIM_SoftwareElementVersionCheck"

'*-----------------------------------------------------------------------------
'* CIM_SoftwareFeature クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMSoftwareFeature = "CIM_SoftwareFeature"

'*-----------------------------------------------------------------------------
'* CIM_SoftwareFeatureSAPImplementation クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMSoftwareFeatureSAPImplementation = "CIM_SoftwareFeatureSAPImplementation"

'*-----------------------------------------------------------------------------
'* CIM_SoftwareFeatureServiceImplementation クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMSoftwareFeatureServiceImplementation = "CIM_SoftwareFeatureServiceImplementation"

'*-----------------------------------------------------------------------------
'* CIM_SoftwareFeatureSoftwareElements クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMSoftwareFeatureSoftwareElements = "CIM_SoftwareFeatureSoftwareElements"

'*-----------------------------------------------------------------------------
'* CIM_SpareGroup クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMSpareGroup = "CIM_SpareGroup"

'*-----------------------------------------------------------------------------
'* CIM_StatisticalInformation クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMStatisticalInformation = "CIM_StatisticalInformation"

'*-----------------------------------------------------------------------------
'* CIM_Statistics クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMStatistics = "CIM_Statistics"

'*-----------------------------------------------------------------------------
'* CIM_StorageDefect クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMStorageDefect = "CIM_StorageDefect"

'*-----------------------------------------------------------------------------
'* CIM_StorageError クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMStorageError = "CIM_StorageError"

'*-----------------------------------------------------------------------------
'* CIM_StorageExtent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMStorageExtent = "CIM_StorageExtent"

'*-----------------------------------------------------------------------------
'* CIM_StorageRedundancyGroup クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMStorageRedundancyGroup = "CIM_StorageRedundancyGroup"

'*-----------------------------------------------------------------------------
'* CIM_SupportAccess クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMSupportAccess = "CIM_SupportAccess"

'*-----------------------------------------------------------------------------
'* CIM_SwapSpaceCheck クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMSwapSpaceCheck = "CIM_SwapSpaceCheck"

'*-----------------------------------------------------------------------------
'* CIM_System クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMSystem = "CIM_System"

'*-----------------------------------------------------------------------------
'* CIM_SystemComponent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMSystemComponent = "CIM_SystemComponent"

'*-----------------------------------------------------------------------------
'* CIM_SystemDevice クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMSystemDevice = "CIM_SystemDevice"

'*-----------------------------------------------------------------------------
'* CIM_SystemResource クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMSystemResource = "CIM_SystemResource"

'*-----------------------------------------------------------------------------
'* CIM_Tachometer クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMTachometer = "CIM_Tachometer"

'*-----------------------------------------------------------------------------
'* CIM_TapeDrive クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMTapeDrive = "CIM_TapeDrive"

'*-----------------------------------------------------------------------------
'* CIM_TemperatureSensor クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMTemperatureSensor = "CIM_TemperatureSensor"

'*-----------------------------------------------------------------------------
'* CIM_Thread クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMThread = "CIM_Thread"

'*-----------------------------------------------------------------------------
'* CIM_ToDirectoryAction クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMToDirectoryAction = "CIM_ToDirectoryAction"

'*-----------------------------------------------------------------------------
'* CIM_ToDirectorySpecification クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMToDirectorySpecification = "CIM_ToDirectorySpecification"

'*-----------------------------------------------------------------------------
'* CIM_UninterruptiblePowerSupply クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMUninterruptiblePowerSupply = "CIM_UninterruptiblePowerSupply"

'*-----------------------------------------------------------------------------
'* CIM_UnitaryComputerSystem クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMUnitaryComputerSystem = "CIM_UnitaryComputerSystem"

'*-----------------------------------------------------------------------------
'* CIM_USBController クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMUSBController = "CIM_USBController"

'*-----------------------------------------------------------------------------
'* CIM_USBControllerHasHub クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMUSBControllerHasHub = "CIM_USBControllerHasHub"

'*-----------------------------------------------------------------------------
'* CIM_UserDevice クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMUserDevice = "CIM_UserDevice"

'*-----------------------------------------------------------------------------
'* CIM_VersionCompatibilityCheck クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMVersionCompatibilityCheck = "CIM_VersionCompatibilityCheck"

'*-----------------------------------------------------------------------------
'* CIM_VideoBIOSElement クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMVideoBIOSElement = "CIM_VideoBIOSElement"

'*-----------------------------------------------------------------------------
'* CIM_VideoBIOSFeature クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMVideoBIOSFeature = "CIM_VideoBIOSFeature"

'*-----------------------------------------------------------------------------
'* CIM_VideoBIOSFeatureVideoBIOSElements クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMVideoBIOSFeatureVideoBIOSElements = "CIM_VideoBIOSFeatureVideoBIOSElements"

'*-----------------------------------------------------------------------------
'* CIM_VideoController クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMVideoController = "CIM_VideoController"

'*-----------------------------------------------------------------------------
'* CIM_VideoControllerResolution クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMVideoControllerResolution = "CIM_VideoControllerResolution"

'*-----------------------------------------------------------------------------
'* CIM_VideoSetting クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMVideoSetting = "CIM_VideoSetting"

'*-----------------------------------------------------------------------------
'* CIM_VolatileStorage クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMVolatileStorage = "CIM_VolatileStorage"

'*-----------------------------------------------------------------------------
'* CIM_VoltageSensor クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMVoltageSensor = "CIM_VoltageSensor"

'*-----------------------------------------------------------------------------
'* CIM_VolumeSet クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMVolumeSet = "CIM_VolumeSet"

'*-----------------------------------------------------------------------------
'* CIM_WORMDrive クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameCIMWORMDrive = "CIM_WORMDrive"

'*-----------------------------------------------------------------------------
'* MSFT_NCProvAccessCheck クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNCProvAccessCheck = "MSFT_NCProvAccessCheck"

'*-----------------------------------------------------------------------------
'* MSFT_NCProvCancelQuery クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNCProvCancelQuery = "MSFT_NCProvCancelQuery"

'*-----------------------------------------------------------------------------
'* MSFT_NCProvClientConnected クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNCProvClientConnected = "MSFT_NCProvClientConnected"

'*-----------------------------------------------------------------------------
'* MSFT_NCProvEvent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNCProvEvent = "MSFT_NCProvEvent"

'*-----------------------------------------------------------------------------
'* MSFT_NCProvNewQuery クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNCProvNewQuery = "MSFT_NCProvNewQuery"

'*-----------------------------------------------------------------------------
'* MSFT_NetBadAccount クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetBadAccount = "MSFT_NetBadAccount"

'*-----------------------------------------------------------------------------
'* MSFT_NetBadServiceState クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetBadServiceState = "MSFT_NetBadServiceState"

'*-----------------------------------------------------------------------------
'* MSFT_NetBootSystemDriversFailed クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetBootSystemDriversFailed = "MSFT_NetBootSystemDriversFailed"

'*-----------------------------------------------------------------------------
'* MSFT_NetCallToFunctionFailed クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetCallToFunctionFailed = "MSFT_NetCallToFunctionFailed"

'*-----------------------------------------------------------------------------
'* MSFT_NetCallToFunctionFailedII クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetCallToFunctionFailedII = "MSFT_NetCallToFunctionFailedII"

'*-----------------------------------------------------------------------------
'* MSFT_NetCircularDependencyAuto クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetCircularDependencyAuto = "MSFT_NetCircularDependencyAuto"

'*-----------------------------------------------------------------------------
'* MSFT_NetCircularDependencyDemand クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetCircularDependencyDemand = "MSFT_NetCircularDependencyDemand"

'*-----------------------------------------------------------------------------
'* MSFT_NetConnectionTimeout クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetConnectionTimeout = "MSFT_NetConnectionTimeout"

'*-----------------------------------------------------------------------------
'* MSFT_NetDependOnLaterGroup クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetDependOnLaterGroup = "MSFT_NetDependOnLaterGroup"

'*-----------------------------------------------------------------------------
'* MSFT_NetDependOnLaterService クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetDependOnLaterService = "MSFT_NetDependOnLaterService"

'*-----------------------------------------------------------------------------
'* MSFT_NetFirstLogonFailed クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetFirstLogonFailed = "MSFT_NetFirstLogonFailed"

'*-----------------------------------------------------------------------------
'* MSFT_NetFirstLogonFailedII クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetFirstLogonFailedII = "MSFT_NetFirstLogonFailedII"

'*-----------------------------------------------------------------------------
'* MSFT_NetReadfileTimeout クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetReadfileTimeout = "MSFT_NetReadfileTimeout"

'*-----------------------------------------------------------------------------
'* MSFT_NetRevertedToLastKnownGood クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetRevertedToLastKnownGood = "MSFT_NetRevertedToLastKnownGood"

'*-----------------------------------------------------------------------------
'* MSFT_NetServiceConfigBackoutFailed クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetServiceConfigBackoutFailed = "MSFT_NetServiceConfigBackoutFailed"

'*-----------------------------------------------------------------------------
'* MSFT_NetServiceControlSuccess クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetServiceControlSuccess = "MSFT_NetServiceControlSuccess"

'*-----------------------------------------------------------------------------
'* MSFT_NetServiceCrash クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetServiceCrash = "MSFT_NetServiceCrash"

'*-----------------------------------------------------------------------------
'* MSFT_NetServiceCrashNoAction クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetServiceCrashNoAction = "MSFT_NetServiceCrashNoAction"

'*-----------------------------------------------------------------------------
'* MSFT_NetServiceDifferentPIDConnected クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetServiceDifferentPIDConnected = "MSFT_NetServiceDifferentPIDConnected"

'*-----------------------------------------------------------------------------
'* MSFT_NetServiceExitFailed クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetServiceExitFailed = "MSFT_NetServiceExitFailed"

'*-----------------------------------------------------------------------------
'* MSFT_NetServiceExitFailedSpecific クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetServiceExitFailedSpecific = "MSFT_NetServiceExitFailedSpecific"

'*-----------------------------------------------------------------------------
'* MSFT_NetServiceLogonTypeNotGranted クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetServiceLogonTypeNotGranted = "MSFT_NetServiceLogonTypeNotGranted"

'*-----------------------------------------------------------------------------
'* MSFT_NetServiceNotInteractive クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetServiceNotInteractive = "MSFT_NetServiceNotInteractive"

'*-----------------------------------------------------------------------------
'* MSFT_NetServiceRecoveryFailed クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetServiceRecoveryFailed = "MSFT_NetServiceRecoveryFailed"

'*-----------------------------------------------------------------------------
'* MSFT_NetServiceShutdownFailed クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetServiceShutdownFailed = "MSFT_NetServiceShutdownFailed"

'*-----------------------------------------------------------------------------
'* MSFT_NetServiceSlowStartup クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetServiceSlowStartup = "MSFT_NetServiceSlowStartup"

'*-----------------------------------------------------------------------------
'* MSFT_NetServiceStartFailed クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetServiceStartFailed = "MSFT_NetServiceStartFailed"

'*-----------------------------------------------------------------------------
'* MSFT_NetServiceStartFailedGroup クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetServiceStartFailedGroup = "MSFT_NetServiceStartFailedGroup"

'*-----------------------------------------------------------------------------
'* MSFT_NetServiceStartFailedII クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetServiceStartFailedII = "MSFT_NetServiceStartFailedII"

'*-----------------------------------------------------------------------------
'* MSFT_NetServiceStartFailedNone クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetServiceStartFailedNone = "MSFT_NetServiceStartFailedNone"

'*-----------------------------------------------------------------------------
'* MSFT_NetServiceStartHung クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetServiceStartHung = "MSFT_NetServiceStartHung"

'*-----------------------------------------------------------------------------
'* MSFT_NetServiceStartTypeChanged クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetServiceStartTypeChanged = "MSFT_NetServiceStartTypeChanged"

'*-----------------------------------------------------------------------------
'* MSFT_NetServiceStatusSuccess クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetServiceStatusSuccess = "MSFT_NetServiceStatusSuccess"

'*-----------------------------------------------------------------------------
'* MSFT_NetServiceStopControlSuccess クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetServiceStopControlSuccess = "MSFT_NetServiceStopControlSuccess"

'*-----------------------------------------------------------------------------
'* MSFT_NetSevereServiceFailed クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetSevereServiceFailed = "MSFT_NetSevereServiceFailed"

'*-----------------------------------------------------------------------------
'* MSFT_NetTakeOwnership クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetTakeOwnership = "MSFT_NetTakeOwnership"

'*-----------------------------------------------------------------------------
'* MSFT_NetTransactInvalid クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetTransactInvalid = "MSFT_NetTransactInvalid"

'*-----------------------------------------------------------------------------
'* MSFT_NetTransactTimeout クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTNetTransactTimeout = "MSFT_NetTransactTimeout"

'*-----------------------------------------------------------------------------
'* Msft_Providers クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftProviders = "Msft_Providers"

'*-----------------------------------------------------------------------------
'* MSFT_SCMEvent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTSCMEvent = "MSFT_SCMEvent"

'*-----------------------------------------------------------------------------
'* MSFT_SCMEventLogEvent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTSCMEventLogEvent = "MSFT_SCMEventLogEvent"

'*-----------------------------------------------------------------------------
'* MSFT_WMI_GenericNonCOMEvent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTWMIGenericNonCOMEvent = "MSFT_WMI_GenericNonCOMEvent"

'*-----------------------------------------------------------------------------
'* MSFT_WmiCancelNotificationSink クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTWmiCancelNotificationSink = "MSFT_WmiCancelNotificationSink"

'*-----------------------------------------------------------------------------
'* MSFT_WmiConsumerProviderEvent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTWmiConsumerProviderEvent = "MSFT_WmiConsumerProviderEvent"

'*-----------------------------------------------------------------------------
'* MSFT_WmiConsumerProviderLoaded クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTWmiConsumerProviderLoaded = "MSFT_WmiConsumerProviderLoaded"

'*-----------------------------------------------------------------------------
'* MSFT_WmiConsumerProviderSinkLoaded クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTWmiConsumerProviderSinkLoaded = "MSFT_WmiConsumerProviderSinkLoaded"

'*-----------------------------------------------------------------------------
'* MSFT_WmiConsumerProviderSinkUnloaded クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTWmiConsumerProviderSinkUnloaded = "MSFT_WmiConsumerProviderSinkUnloaded"

'*-----------------------------------------------------------------------------
'* MSFT_WmiConsumerProviderUnloaded クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTWmiConsumerProviderUnloaded = "MSFT_WmiConsumerProviderUnloaded"

'*-----------------------------------------------------------------------------
'* MSFT_WmiEssEvent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTWmiEssEvent = "MSFT_WmiEssEvent"

'*-----------------------------------------------------------------------------
'* MSFT_WmiFilterActivated クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTWmiFilterActivated = "MSFT_WmiFilterActivated"

'*-----------------------------------------------------------------------------
'* MSFT_WmiFilterDeactivated クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTWmiFilterDeactivated = "MSFT_WmiFilterDeactivated"

'*-----------------------------------------------------------------------------
'* MSFT_WmiFilterEvent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTWmiFilterEvent = "MSFT_WmiFilterEvent"

'*-----------------------------------------------------------------------------
'* Msft_WmiProvider_AccessCheck_Post クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftWmiProviderAccessCheckPost = "Msft_WmiProvider_AccessCheck_Post"

'*-----------------------------------------------------------------------------
'* Msft_WmiProvider_AccessCheck_Pre クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftWmiProviderAccessCheckPre = "Msft_WmiProvider_AccessCheck_Pre"

'*-----------------------------------------------------------------------------
'* Msft_WmiProvider_CancelQuery_Post クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftWmiProviderCancelQueryPost = "Msft_WmiProvider_CancelQuery_Post"

'*-----------------------------------------------------------------------------
'* Msft_WmiProvider_CancelQuery_Pre クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftWmiProviderCancelQueryPre = "Msft_WmiProvider_CancelQuery_Pre"

'*-----------------------------------------------------------------------------
'* Msft_WmiProvider_ComServerLoadOperationEvent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftWmiProviderComServerLoadOperationEvent = "Msft_WmiProvider_ComServerLoadOperationEvent"

'*-----------------------------------------------------------------------------
'* Msft_WmiProvider_ComServerLoadOperationFailureEvent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftWmiProviderComServerLoadOperationFailureEvent = "Msft_WmiProvider_ComServerLoadOperationFailureEvent"

'*-----------------------------------------------------------------------------
'* Msft_WmiProvider_Counters クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftWmiProviderCounters = "Msft_WmiProvider_Counters"

'*-----------------------------------------------------------------------------
'* Msft_WmiProvider_CreateClassEnumAsyncEvent_Post クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftWmiProviderCreateClassEnumAsyncEventPost = "Msft_WmiProvider_CreateClassEnumAsyncEvent_Post"

'*-----------------------------------------------------------------------------
'* Msft_WmiProvider_CreateClassEnumAsyncEvent_Pre クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftWmiProviderCreateClassEnumAsyncEventPre = "Msft_WmiProvider_CreateClassEnumAsyncEvent_Pre"

'*-----------------------------------------------------------------------------
'* Msft_WmiProvider_CreateInstanceEnumAsyncEvent_Post クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftWmiProviderCreateInstanceEnumAsyncEventPost = "Msft_WmiProvider_CreateInstanceEnumAsyncEvent_Post"

'*-----------------------------------------------------------------------------
'* Msft_WmiProvider_CreateInstanceEnumAsyncEvent_Pre クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftWmiProviderCreateInstanceEnumAsyncEventPre = "Msft_WmiProvider_CreateInstanceEnumAsyncEvent_Pre"

'*-----------------------------------------------------------------------------
'* Msft_WmiProvider_DeleteClassAsyncEvent_Post クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftWmiProviderDeleteClassAsyncEventPost = "Msft_WmiProvider_DeleteClassAsyncEvent_Post"

'*-----------------------------------------------------------------------------
'* Msft_WmiProvider_DeleteClassAsyncEvent_Pre クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftWmiProviderDeleteClassAsyncEventPre = "Msft_WmiProvider_DeleteClassAsyncEvent_Pre"

'*-----------------------------------------------------------------------------
'* Msft_WmiProvider_DeleteInstanceAsyncEvent_Post クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftWmiProviderDeleteInstanceAsyncEventPost = "Msft_WmiProvider_DeleteInstanceAsyncEvent_Post"

'*-----------------------------------------------------------------------------
'* Msft_WmiProvider_DeleteInstanceAsyncEvent_Pre クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftWmiProviderDeleteInstanceAsyncEventPre = "Msft_WmiProvider_DeleteInstanceAsyncEvent_Pre"

'*-----------------------------------------------------------------------------
'* Msft_WmiProvider_ExecMethodAsyncEvent_Post クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftWmiProviderExecMethodAsyncEventPost = "Msft_WmiProvider_ExecMethodAsyncEvent_Post"

'*-----------------------------------------------------------------------------
'* Msft_WmiProvider_ExecMethodAsyncEvent_Pre クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftWmiProviderExecMethodAsyncEventPre = "Msft_WmiProvider_ExecMethodAsyncEvent_Pre"

'*-----------------------------------------------------------------------------
'* Msft_WmiProvider_ExecQueryAsyncEvent_Post クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftWmiProviderExecQueryAsyncEventPost = "Msft_WmiProvider_ExecQueryAsyncEvent_Post"

'*-----------------------------------------------------------------------------
'* Msft_WmiProvider_ExecQueryAsyncEvent_Pre クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftWmiProviderExecQueryAsyncEventPre = "Msft_WmiProvider_ExecQueryAsyncEvent_Pre"

'*-----------------------------------------------------------------------------
'* Msft_WmiProvider_GetObjectAsyncEvent_Post クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftWmiProviderGetObjectAsyncEventPost = "Msft_WmiProvider_GetObjectAsyncEvent_Post"

'*-----------------------------------------------------------------------------
'* Msft_WmiProvider_GetObjectAsyncEvent_Pre クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftWmiProviderGetObjectAsyncEventPre = "Msft_WmiProvider_GetObjectAsyncEvent_Pre"

'*-----------------------------------------------------------------------------
'* Msft_WmiProvider_InitializationOperationEvent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftWmiProviderInitializationOperationEvent = "Msft_WmiProvider_InitializationOperationEvent"

'*-----------------------------------------------------------------------------
'* Msft_WmiProvider_InitializationOperationFailureEvent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftWmiProviderInitializationOperationFailureEvent = "Msft_WmiProvider_InitializationOperationFailureEvent"

'*-----------------------------------------------------------------------------
'* Msft_WmiProvider_LoadOperationEvent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftWmiProviderLoadOperationEvent = "Msft_WmiProvider_LoadOperationEvent"

'*-----------------------------------------------------------------------------
'* Msft_WmiProvider_LoadOperationFailureEvent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftWmiProviderLoadOperationFailureEvent = "Msft_WmiProvider_LoadOperationFailureEvent"

'*-----------------------------------------------------------------------------
'* Msft_WmiProvider_NewQuery_Post クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftWmiProviderNewQueryPost = "Msft_WmiProvider_NewQuery_Post"

'*-----------------------------------------------------------------------------
'* Msft_WmiProvider_NewQuery_Pre クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftWmiProviderNewQueryPre = "Msft_WmiProvider_NewQuery_Pre"

'*-----------------------------------------------------------------------------
'* Msft_WmiProvider_OperationEvent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftWmiProviderOperationEvent = "Msft_WmiProvider_OperationEvent"

'*-----------------------------------------------------------------------------
'* Msft_WmiProvider_OperationEvent_Post クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftWmiProviderOperationEventPost = "Msft_WmiProvider_OperationEvent_Post"

'*-----------------------------------------------------------------------------
'* Msft_WmiProvider_OperationEvent_Pre クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftWmiProviderOperationEventPre = "Msft_WmiProvider_OperationEvent_Pre"

'*-----------------------------------------------------------------------------
'* Msft_WmiProvider_ProvideEvents_Post クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftWmiProviderProvideEventsPost = "Msft_WmiProvider_ProvideEvents_Post"

'*-----------------------------------------------------------------------------
'* Msft_WmiProvider_ProvideEvents_Pre クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftWmiProviderProvideEventsPre = "Msft_WmiProvider_ProvideEvents_Pre"

'*-----------------------------------------------------------------------------
'* Msft_WmiProvider_PutClassAsyncEvent_Post クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftWmiProviderPutClassAsyncEventPost = "Msft_WmiProvider_PutClassAsyncEvent_Post"

'*-----------------------------------------------------------------------------
'* Msft_WmiProvider_PutClassAsyncEvent_Pre クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftWmiProviderPutClassAsyncEventPre = "Msft_WmiProvider_PutClassAsyncEvent_Pre"

'*-----------------------------------------------------------------------------
'* Msft_WmiProvider_PutInstanceAsyncEvent_Post クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftWmiProviderPutInstanceAsyncEventPost = "Msft_WmiProvider_PutInstanceAsyncEvent_Post"

'*-----------------------------------------------------------------------------
'* Msft_WmiProvider_PutInstanceAsyncEvent_Pre クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftWmiProviderPutInstanceAsyncEventPre = "Msft_WmiProvider_PutInstanceAsyncEvent_Pre"

'*-----------------------------------------------------------------------------
'* Msft_WmiProvider_UnLoadOperationEvent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMsftWmiProviderUnLoadOperationEvent = "Msft_WmiProvider_UnLoadOperationEvent"

'*-----------------------------------------------------------------------------
'* MSFT_WmiProviderEvent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTWmiProviderEvent = "MSFT_WmiProviderEvent"

'*-----------------------------------------------------------------------------
'* MSFT_WmiRegisterNotificationSink クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTWmiRegisterNotificationSink = "MSFT_WmiRegisterNotificationSink"

'*-----------------------------------------------------------------------------
'* MSFT_WmiSelfEvent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTWmiSelfEvent = "MSFT_WmiSelfEvent"

'*-----------------------------------------------------------------------------
'* MSFT_WmiThreadPoolEvent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTWmiThreadPoolEvent = "MSFT_WmiThreadPoolEvent"

'*-----------------------------------------------------------------------------
'* MSFT_WmiThreadPoolThreadCreated クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTWmiThreadPoolThreadCreated = "MSFT_WmiThreadPoolThreadCreated"

'*-----------------------------------------------------------------------------
'* MSFT_WmiThreadPoolThreadDeleted クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameMSFTWmiThreadPoolThreadDeleted = "MSFT_WmiThreadPoolThreadDeleted"

'*-----------------------------------------------------------------------------
'* Win32_1394Controller クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin321394Controller = "Win32_1394Controller"

'*-----------------------------------------------------------------------------
'* Win32_1394ControllerDevice クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin321394ControllerDevice = "Win32_1394ControllerDevice"

'*-----------------------------------------------------------------------------
'* Win32_Account クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32Account = "Win32_Account"

'*-----------------------------------------------------------------------------
'* Win32_AccountSID クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32AccountSID = "Win32_AccountSID"

'*-----------------------------------------------------------------------------
'* Win32_ACE クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ACE = "Win32_ACE"

'*-----------------------------------------------------------------------------
'* Win32_ActionCheck クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ActionCheck = "Win32_ActionCheck"

'*-----------------------------------------------------------------------------
'* Win32_ActiveRoute クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ActiveRoute = "Win32_ActiveRoute"

'*-----------------------------------------------------------------------------
'* Win32_AllocatedResource クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32AllocatedResource = "Win32_AllocatedResource"

'*-----------------------------------------------------------------------------
'* Win32_ApplicationCommandLine クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ApplicationCommandLine = "Win32_ApplicationCommandLine"

'*-----------------------------------------------------------------------------
'* Win32_ApplicationService クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ApplicationService = "Win32_ApplicationService"

'*-----------------------------------------------------------------------------
'* Win32_AssociatedProcessorMemory クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32AssociatedProcessorMemory = "Win32_AssociatedProcessorMemory"

'*-----------------------------------------------------------------------------
'* Win32_AutochkSetting クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32AutochkSetting = "Win32_AutochkSetting"

'*-----------------------------------------------------------------------------
'* Win32_BaseBoard クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32BaseBoard = "Win32_BaseBoard"

'*-----------------------------------------------------------------------------
'* Win32_BaseService クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32BaseService = "Win32_BaseService"

'*-----------------------------------------------------------------------------
'* Win32_Battery クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32Battery = "Win32_Battery"

'*-----------------------------------------------------------------------------
'* Win32_Binary クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32Binary = "Win32_Binary"

'*-----------------------------------------------------------------------------
'* Win32_BindImageAction クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32BindImageAction = "Win32_BindImageAction"

'*-----------------------------------------------------------------------------
'* Win32_BootConfiguration クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32BootConfiguration = "Win32_BootConfiguration"

'*-----------------------------------------------------------------------------
'* Win32_Bus クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32Bus = "Win32_Bus"

'*-----------------------------------------------------------------------------
'* Win32_CacheMemory クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32CacheMemory = "Win32_CacheMemory"

'*-----------------------------------------------------------------------------
'* Win32_CDROMDrive クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32CDROMDrive = "Win32_CDROMDrive"

'*-----------------------------------------------------------------------------
'* Win32_CheckCheck クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32CheckCheck = "Win32_CheckCheck"

'*-----------------------------------------------------------------------------
'* Win32_CIMLogicalDeviceCIMDataFile クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32CIMLogicalDeviceCIMDataFile = "Win32_CIMLogicalDeviceCIMDataFile"

'*-----------------------------------------------------------------------------
'* Win32_ClassicCOMApplicationClasses クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ClassicCOMApplicationClasses = "Win32_ClassicCOMApplicationClasses"

'*-----------------------------------------------------------------------------
'* Win32_ClassicCOMClass クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ClassicCOMClass = "Win32_ClassicCOMClass"

'*-----------------------------------------------------------------------------
'* Win32_ClassicCOMClassSetting クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ClassicCOMClassSetting = "Win32_ClassicCOMClassSetting"

'*-----------------------------------------------------------------------------
'* Win32_ClassicCOMClassSettings クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ClassicCOMClassSettings = "Win32_ClassicCOMClassSettings"

'*-----------------------------------------------------------------------------
'* Win32_ClassInfoAction クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ClassInfoAction = "Win32_ClassInfoAction"

'*-----------------------------------------------------------------------------
'* Win32_ClientApplicationSetting クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ClientApplicationSetting = "Win32_ClientApplicationSetting"

'*-----------------------------------------------------------------------------
'* Win32_ClusterShare クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ClusterShare = "Win32_ClusterShare"

'*-----------------------------------------------------------------------------
'* Win32_CodecFile クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32CodecFile = "Win32_CodecFile"

'*-----------------------------------------------------------------------------
'* Win32_CollectionStatistics クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32CollectionStatistics = "Win32_CollectionStatistics"

'*-----------------------------------------------------------------------------
'* Win32_COMApplication クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32COMApplication = "Win32_COMApplication"

'*-----------------------------------------------------------------------------
'* Win32_COMApplicationClasses クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32COMApplicationClasses = "Win32_COMApplicationClasses"

'*-----------------------------------------------------------------------------
'* Win32_COMApplicationSettings クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32COMApplicationSettings = "Win32_COMApplicationSettings"

'*-----------------------------------------------------------------------------
'* Win32_COMClass クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32COMClass = "Win32_COMClass"

'*-----------------------------------------------------------------------------
'* Win32_ComClassAutoEmulator クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ComClassAutoEmulator = "Win32_ComClassAutoEmulator"

'*-----------------------------------------------------------------------------
'* Win32_ComClassEmulator クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ComClassEmulator = "Win32_ComClassEmulator"

'*-----------------------------------------------------------------------------
'* Win32_CommandLineAccess クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32CommandLineAccess = "Win32_CommandLineAccess"

'*-----------------------------------------------------------------------------
'* Win32_ComponentCategory クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ComponentCategory = "Win32_ComponentCategory"

'*-----------------------------------------------------------------------------
'* Win32_ComputerShutdownEvent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ComputerShutdownEvent = "Win32_ComputerShutdownEvent"

'*-----------------------------------------------------------------------------
'* Win32_ComputerSystemEvent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ComputerSystemEvent = "Win32_ComputerSystemEvent"

'*-----------------------------------------------------------------------------
'* Win32_ComputerSystemProcessor クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ComputerSystemProcessor = "Win32_ComputerSystemProcessor"

'*-----------------------------------------------------------------------------
'* Win32_COMSetting クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32COMSetting = "Win32_COMSetting"

'*-----------------------------------------------------------------------------
'* Win32_Condition クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32Condition = "Win32_Condition"

'*-----------------------------------------------------------------------------
'* Win32_ConnectionShare クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ConnectionShare = "Win32_ConnectionShare"

'*-----------------------------------------------------------------------------
'* Win32_ControllerHasHub クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ControllerHasHub = "Win32_ControllerHasHub"

'*-----------------------------------------------------------------------------
'* Win32_CreateFolderAction クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32CreateFolderAction = "Win32_CreateFolderAction"

'*-----------------------------------------------------------------------------
'* Win32_CurrentProbe クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32CurrentProbe = "Win32_CurrentProbe"

'*-----------------------------------------------------------------------------
'* Win32_CurrentTime クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32CurrentTime = "Win32_CurrentTime"

'*-----------------------------------------------------------------------------
'* Win32_DCOMApplication クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32DCOMApplication = "Win32_DCOMApplication"

'*-----------------------------------------------------------------------------
'* Win32_DCOMApplicationAccessAllowedSetting クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32DCOMApplicationAccessAllowedSetting = "Win32_DCOMApplicationAccessAllowedSetting"

'*-----------------------------------------------------------------------------
'* Win32_DCOMApplicationLaunchAllowedSetting クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32DCOMApplicationLaunchAllowedSetting = "Win32_DCOMApplicationLaunchAllowedSetting"

'*-----------------------------------------------------------------------------
'* Win32_DCOMApplicationSetting クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32DCOMApplicationSetting = "Win32_DCOMApplicationSetting"

'*-----------------------------------------------------------------------------
'* Win32_DependentService クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32DependentService = "Win32_DependentService"

'*-----------------------------------------------------------------------------
'* Win32_Desktop クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32Desktop = "Win32_Desktop"

'*-----------------------------------------------------------------------------
'* Win32_DesktopMonitor クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32DesktopMonitor = "Win32_DesktopMonitor"

'*-----------------------------------------------------------------------------
'* Win32_DeviceBus クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32DeviceBus = "Win32_DeviceBus"

'*-----------------------------------------------------------------------------
'* Win32_DeviceChangeEvent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32DeviceChangeEvent = "Win32_DeviceChangeEvent"

'*-----------------------------------------------------------------------------
'* Win32_DeviceMemoryAddress クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32DeviceMemoryAddress = "Win32_DeviceMemoryAddress"

'*-----------------------------------------------------------------------------
'* Win32_DeviceSettings クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32DeviceSettings = "Win32_DeviceSettings"

'*-----------------------------------------------------------------------------
'* Win32_DfsNode クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32DfsNode = "Win32_DfsNode"

'*-----------------------------------------------------------------------------
'* Win32_DfsNodeTarget クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32DfsNodeTarget = "Win32_DfsNodeTarget"

'*-----------------------------------------------------------------------------
'* Win32_DfsTarget クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32DfsTarget = "Win32_DfsTarget"

'*-----------------------------------------------------------------------------
'* Win32_Directory クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32Directory = "Win32_Directory"

'*-----------------------------------------------------------------------------
'* Win32_DirectorySpecification クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32DirectorySpecification = "Win32_DirectorySpecification"

'*-----------------------------------------------------------------------------
'* Win32_DiskDrivePhysicalMedia クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32DiskDrivePhysicalMedia = "Win32_DiskDrivePhysicalMedia"

'*-----------------------------------------------------------------------------
'* Win32_DiskDriveToDiskPartition クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32DiskDriveToDiskPartition = "Win32_DiskDriveToDiskPartition"

'*-----------------------------------------------------------------------------
'* Win32_DiskPartition クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32DiskPartition = "Win32_DiskPartition"

'*-----------------------------------------------------------------------------
'* Win32_DiskQuota クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32DiskQuota = "Win32_DiskQuota"

'*-----------------------------------------------------------------------------
'* Win32_DisplayConfiguration クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32DisplayConfiguration = "Win32_DisplayConfiguration"

'*-----------------------------------------------------------------------------
'* Win32_DisplayControllerConfiguration クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32DisplayControllerConfiguration = "Win32_DisplayControllerConfiguration"

'*-----------------------------------------------------------------------------
'* Win32_DMAChannel クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32DMAChannel = "Win32_DMAChannel"

'*-----------------------------------------------------------------------------
'* Win32_DriverForDevice クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32DriverForDevice = "Win32_DriverForDevice"

'*-----------------------------------------------------------------------------
'* Win32_DuplicateFileAction クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32DuplicateFileAction = "Win32_DuplicateFileAction"

'*-----------------------------------------------------------------------------
'* Win32_Environment クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32Environment = "Win32_Environment"

'*-----------------------------------------------------------------------------
'* Win32_EnvironmentSpecification クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32EnvironmentSpecification = "Win32_EnvironmentSpecification"

'*-----------------------------------------------------------------------------
'* Win32_ExtensionInfoAction クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ExtensionInfoAction = "Win32_ExtensionInfoAction"

'*-----------------------------------------------------------------------------
'* Win32_Fan クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32Fan = "Win32_Fan"

'*-----------------------------------------------------------------------------
'* Win32_FileSpecification クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32FileSpecification = "Win32_FileSpecification"

'*-----------------------------------------------------------------------------
'* Win32_FloppyController クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32FloppyController = "Win32_FloppyController"

'*-----------------------------------------------------------------------------
'* Win32_FloppyDrive クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32FloppyDrive = "Win32_FloppyDrive"

'*-----------------------------------------------------------------------------
'* Win32_FolderRedirection クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32FolderRedirection = "Win32_FolderRedirection"

'*-----------------------------------------------------------------------------
'* Win32_FolderRedirectionHealth クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32FolderRedirectionHealth = "Win32_FolderRedirectionHealth"

'*-----------------------------------------------------------------------------
'* Win32_FolderRedirectionHealthConfiguration クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32FolderRedirectionHealthConfiguration = "Win32_FolderRedirectionHealthConfiguration"

'*-----------------------------------------------------------------------------
'* Win32_FolderRedirectionUserConfiguration クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32FolderRedirectionUserConfiguration = "Win32_FolderRedirectionUserConfiguration"

'*-----------------------------------------------------------------------------
'* Win32_FontInfoAction クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32FontInfoAction = "Win32_FontInfoAction"

'*-----------------------------------------------------------------------------
'* Win32_Group クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32Group = "Win32_Group"

'*-----------------------------------------------------------------------------
'* Win32_GroupInDomain クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32GroupInDomain = "Win32_GroupInDomain"

'*-----------------------------------------------------------------------------
'* Win32_GroupUser クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32GroupUser = "Win32_GroupUser"

'*-----------------------------------------------------------------------------
'* Win32_HeatPipe クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32HeatPipe = "Win32_HeatPipe"

'*-----------------------------------------------------------------------------
'* Win32_IDEController クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32IDEController = "Win32_IDEController"

'*-----------------------------------------------------------------------------
'* Win32_IDEControllerDevice クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32IDEControllerDevice = "Win32_IDEControllerDevice"

'*-----------------------------------------------------------------------------
'* Win32_ImplementedCategory クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ImplementedCategory = "Win32_ImplementedCategory"

'*-----------------------------------------------------------------------------
'* Win32_InfraredDevice クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32InfraredDevice = "Win32_InfraredDevice"

'*-----------------------------------------------------------------------------
'* Win32_IniFileSpecification クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32IniFileSpecification = "Win32_IniFileSpecification"

'*-----------------------------------------------------------------------------
'* Win32_InstalledProgramFramework クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32InstalledProgramFramework = "Win32_InstalledProgramFramework"

'*-----------------------------------------------------------------------------
'* Win32_InstalledSoftwareElement クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32InstalledSoftwareElement = "Win32_InstalledSoftwareElement"

'*-----------------------------------------------------------------------------
'* Win32_InstalledStoreProgram クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32InstalledStoreProgram = "Win32_InstalledStoreProgram"

'*-----------------------------------------------------------------------------
'* Win32_InstalledWin32Program クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32InstalledWin32Program = "Win32_InstalledWin32Program"

'*-----------------------------------------------------------------------------
'* Win32_IP4PersistedRouteTable クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32IP4PersistedRouteTable = "Win32_IP4PersistedRouteTable"

'*-----------------------------------------------------------------------------
'* Win32_IP4RouteTable クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32IP4RouteTable = "Win32_IP4RouteTable"

'*-----------------------------------------------------------------------------
'* Win32_IP4RouteTableEvent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32IP4RouteTableEvent = "Win32_IP4RouteTableEvent"

'*-----------------------------------------------------------------------------
'* Win32_IRQResource クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32IRQResource = "Win32_IRQResource"

'*-----------------------------------------------------------------------------
'* Win32_JobObjectStatus クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32JobObjectStatus = "Win32_JobObjectStatus"

'*-----------------------------------------------------------------------------
'* Win32_Keyboard クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32Keyboard = "Win32_Keyboard"

'*-----------------------------------------------------------------------------
'* Win32_LaunchCondition クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32LaunchCondition = "Win32_LaunchCondition"

'*-----------------------------------------------------------------------------
'* Win32_LoadOrderGroup クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32LoadOrderGroup = "Win32_LoadOrderGroup"

'*-----------------------------------------------------------------------------
'* Win32_LoadOrderGroupServiceDependencies クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32LoadOrderGroupServiceDependencies = "Win32_LoadOrderGroupServiceDependencies"

'*-----------------------------------------------------------------------------
'* Win32_LoadOrderGroupServiceMembers クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32LoadOrderGroupServiceMembers = "Win32_LoadOrderGroupServiceMembers"

'*-----------------------------------------------------------------------------
'* Win32_LocalTime クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32LocalTime = "Win32_LocalTime"

'*-----------------------------------------------------------------------------
'* Win32_LoggedOnUser クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32LoggedOnUser = "Win32_LoggedOnUser"

'*-----------------------------------------------------------------------------
'* Win32_LogicalDiskRootDirectory クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32LogicalDiskRootDirectory = "Win32_LogicalDiskRootDirectory"

'*-----------------------------------------------------------------------------
'* Win32_LogicalDiskToPartition クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32LogicalDiskToPartition = "Win32_LogicalDiskToPartition"

'*-----------------------------------------------------------------------------
'* Win32_LogicalFileAccess クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32LogicalFileAccess = "Win32_LogicalFileAccess"

'*-----------------------------------------------------------------------------
'* Win32_LogicalFileAuditing クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32LogicalFileAuditing = "Win32_LogicalFileAuditing"

'*-----------------------------------------------------------------------------
'* Win32_LogicalFileGroup クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32LogicalFileGroup = "Win32_LogicalFileGroup"

'*-----------------------------------------------------------------------------
'* Win32_LogicalFileOwner クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32LogicalFileOwner = "Win32_LogicalFileOwner"

'*-----------------------------------------------------------------------------
'* Win32_LogicalFileSecuritySetting クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32LogicalFileSecuritySetting = "Win32_LogicalFileSecuritySetting"

'*-----------------------------------------------------------------------------
'* Win32_LogicalProgramGroup クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32LogicalProgramGroup = "Win32_LogicalProgramGroup"

'*-----------------------------------------------------------------------------
'* Win32_LogicalProgramGroupDirectory クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32LogicalProgramGroupDirectory = "Win32_LogicalProgramGroupDirectory"

'*-----------------------------------------------------------------------------
'* Win32_LogicalProgramGroupItem クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32LogicalProgramGroupItem = "Win32_LogicalProgramGroupItem"

'*-----------------------------------------------------------------------------
'* Win32_LogicalProgramGroupItemDataFile クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32LogicalProgramGroupItemDataFile = "Win32_LogicalProgramGroupItemDataFile"

'*-----------------------------------------------------------------------------
'* Win32_LogicalShareAccess クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32LogicalShareAccess = "Win32_LogicalShareAccess"

'*-----------------------------------------------------------------------------
'* Win32_LogicalShareAuditing クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32LogicalShareAuditing = "Win32_LogicalShareAuditing"

'*-----------------------------------------------------------------------------
'* Win32_LogicalShareSecuritySetting クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32LogicalShareSecuritySetting = "Win32_LogicalShareSecuritySetting"

'*-----------------------------------------------------------------------------
'* Win32_LogonSession クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32LogonSession = "Win32_LogonSession"

'*-----------------------------------------------------------------------------
'* Win32_LogonSessionMappedDisk クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32LogonSessionMappedDisk = "Win32_LogonSessionMappedDisk"

'*-----------------------------------------------------------------------------
'* Win32_LUID クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32LUID = "Win32_LUID"

'*-----------------------------------------------------------------------------
'* Win32_LUIDandAttributes クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32LUIDandAttributes = "Win32_LUIDandAttributes"

'*-----------------------------------------------------------------------------
'* Win32_ManagedSystemElementResource クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ManagedSystemElementResource = "Win32_ManagedSystemElementResource"

'*-----------------------------------------------------------------------------
'* Win32_MappedLogicalDisk クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32MappedLogicalDisk = "Win32_MappedLogicalDisk"

'*-----------------------------------------------------------------------------
'* Win32_MemoryArray クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32MemoryArray = "Win32_MemoryArray"

'*-----------------------------------------------------------------------------
'* Win32_MemoryArrayLocation クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32MemoryArrayLocation = "Win32_MemoryArrayLocation"

'*-----------------------------------------------------------------------------
'* Win32_MemoryDevice クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32MemoryDevice = "Win32_MemoryDevice"

'*-----------------------------------------------------------------------------
'* Win32_MemoryDeviceArray クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32MemoryDeviceArray = "Win32_MemoryDeviceArray"

'*-----------------------------------------------------------------------------
'* Win32_MemoryDeviceLocation クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32MemoryDeviceLocation = "Win32_MemoryDeviceLocation"

'*-----------------------------------------------------------------------------
'* Win32_MethodParameterClass クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32MethodParameterClass = "Win32_MethodParameterClass"

'*-----------------------------------------------------------------------------
'* Win32_MIMEInfoAction クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32MIMEInfoAction = "Win32_MIMEInfoAction"

'*-----------------------------------------------------------------------------
'* Win32_ModuleLoadTrace クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ModuleLoadTrace = "Win32_ModuleLoadTrace"

'*-----------------------------------------------------------------------------
'* Win32_ModuleTrace クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ModuleTrace = "Win32_ModuleTrace"

'*-----------------------------------------------------------------------------
'* Win32_MotherboardDevice クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32MotherboardDevice = "Win32_MotherboardDevice"

'*-----------------------------------------------------------------------------
'* Win32_MountPoint クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32MountPoint = "Win32_MountPoint"

'*-----------------------------------------------------------------------------
'* Win32_MoveFileAction クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32MoveFileAction = "Win32_MoveFileAction"

'*-----------------------------------------------------------------------------
'* Win32_MSIResource クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32MSIResource = "Win32_MSIResource"

'*-----------------------------------------------------------------------------
'* Win32_NamedJobObject クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32NamedJobObject = "Win32_NamedJobObject"

'*-----------------------------------------------------------------------------
'* Win32_NamedJobObjectActgInfo クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32NamedJobObjectActgInfo = "Win32_NamedJobObjectActgInfo"

'*-----------------------------------------------------------------------------
'* Win32_NamedJobObjectLimit クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32NamedJobObjectLimit = "Win32_NamedJobObjectLimit"

'*-----------------------------------------------------------------------------
'* Win32_NamedJobObjectLimitSetting クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32NamedJobObjectLimitSetting = "Win32_NamedJobObjectLimitSetting"

'*-----------------------------------------------------------------------------
'* Win32_NamedJobObjectProcess クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32NamedJobObjectProcess = "Win32_NamedJobObjectProcess"

'*-----------------------------------------------------------------------------
'* Win32_NamedJobObjectSecLimit クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32NamedJobObjectSecLimit = "Win32_NamedJobObjectSecLimit"

'*-----------------------------------------------------------------------------
'* Win32_NamedJobObjectSecLimitSetting クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32NamedJobObjectSecLimitSetting = "Win32_NamedJobObjectSecLimitSetting"

'*-----------------------------------------------------------------------------
'* Win32_NamedJobObjectStatistics クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32NamedJobObjectStatistics = "Win32_NamedJobObjectStatistics"

'*-----------------------------------------------------------------------------
'* Win32_NetworkAdapterSetting クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32NetworkAdapterSetting = "Win32_NetworkAdapterSetting"

'*-----------------------------------------------------------------------------
'* Win32_NetworkClient クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32NetworkClient = "Win32_NetworkClient"

'*-----------------------------------------------------------------------------
'* Win32_NetworkConnection クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32NetworkConnection = "Win32_NetworkConnection"

'*-----------------------------------------------------------------------------
'* Win32_NetworkLoginProfile クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32NetworkLoginProfile = "Win32_NetworkLoginProfile"

'*-----------------------------------------------------------------------------
'* Win32_NetworkProtocol クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32NetworkProtocol = "Win32_NetworkProtocol"

'*-----------------------------------------------------------------------------
'* Win32_NTDomain クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32NTDomain = "Win32_NTDomain"

'*-----------------------------------------------------------------------------
'* Win32_NTEventlogFile クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32NTEventlogFile = "Win32_NTEventlogFile"

'*-----------------------------------------------------------------------------
'* Win32_NTLogEventComputer クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32NTLogEventComputer = "Win32_NTLogEventComputer"

'*-----------------------------------------------------------------------------
'* Win32_NTLogEventLog クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32NTLogEventLog = "Win32_NTLogEventLog"

'*-----------------------------------------------------------------------------
'* Win32_NTLogEventUser クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32NTLogEventUser = "Win32_NTLogEventUser"

'*-----------------------------------------------------------------------------
'* Win32_ODBCAttribute クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ODBCAttribute = "Win32_ODBCAttribute"

'*-----------------------------------------------------------------------------
'* Win32_ODBCDataSourceAttribute クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ODBCDataSourceAttribute = "Win32_ODBCDataSourceAttribute"

'*-----------------------------------------------------------------------------
'* Win32_ODBCDataSourceSpecification クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ODBCDataSourceSpecification = "Win32_ODBCDataSourceSpecification"

'*-----------------------------------------------------------------------------
'* Win32_ODBCDriverAttribute クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ODBCDriverAttribute = "Win32_ODBCDriverAttribute"

'*-----------------------------------------------------------------------------
'* Win32_ODBCDriverSoftwareElement クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ODBCDriverSoftwareElement = "Win32_ODBCDriverSoftwareElement"

'*-----------------------------------------------------------------------------
'* Win32_ODBCDriverSpecification クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ODBCDriverSpecification = "Win32_ODBCDriverSpecification"

'*-----------------------------------------------------------------------------
'* Win32_ODBCSourceAttribute クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ODBCSourceAttribute = "Win32_ODBCSourceAttribute"

'*-----------------------------------------------------------------------------
'* Win32_ODBCTranslatorSpecification クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ODBCTranslatorSpecification = "Win32_ODBCTranslatorSpecification"

'*-----------------------------------------------------------------------------
'* Win32_OfflineFilesAssociatedItems クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32OfflineFilesAssociatedItems = "Win32_OfflineFilesAssociatedItems"

'*-----------------------------------------------------------------------------
'* Win32_OfflineFilesBackgroundSync クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32OfflineFilesBackgroundSync = "Win32_OfflineFilesBackgroundSync"

'*-----------------------------------------------------------------------------
'* Win32_OfflineFilesCache クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32OfflineFilesCache = "Win32_OfflineFilesCache"

'*-----------------------------------------------------------------------------
'* Win32_OfflineFilesChangeInfo クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32OfflineFilesChangeInfo = "Win32_OfflineFilesChangeInfo"

'*-----------------------------------------------------------------------------
'* Win32_OfflineFilesConnectionInfo クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32OfflineFilesConnectionInfo = "Win32_OfflineFilesConnectionInfo"

'*-----------------------------------------------------------------------------
'* Win32_OfflineFilesDirtyInfo クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32OfflineFilesDirtyInfo = "Win32_OfflineFilesDirtyInfo"

'*-----------------------------------------------------------------------------
'* Win32_OfflineFilesDiskSpaceLimit クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32OfflineFilesDiskSpaceLimit = "Win32_OfflineFilesDiskSpaceLimit"

'*-----------------------------------------------------------------------------
'* Win32_OfflineFilesFileSysInfo クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32OfflineFilesFileSysInfo = "Win32_OfflineFilesFileSysInfo"

'*-----------------------------------------------------------------------------
'* Win32_OfflineFilesHealth クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32OfflineFilesHealth = "Win32_OfflineFilesHealth"

'*-----------------------------------------------------------------------------
'* Win32_OfflineFilesItem クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32OfflineFilesItem = "Win32_OfflineFilesItem"

'*-----------------------------------------------------------------------------
'* Win32_OfflineFilesMachineConfiguration クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32OfflineFilesMachineConfiguration = "Win32_OfflineFilesMachineConfiguration"

'*-----------------------------------------------------------------------------
'* Win32_OfflineFilesPinInfo クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32OfflineFilesPinInfo = "Win32_OfflineFilesPinInfo"

'*-----------------------------------------------------------------------------
'* Win32_OfflineFilesSuspendInfo クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32OfflineFilesSuspendInfo = "Win32_OfflineFilesSuspendInfo"

'*-----------------------------------------------------------------------------
'* Win32_OfflineFilesUserConfiguration クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32OfflineFilesUserConfiguration = "Win32_OfflineFilesUserConfiguration"

'*-----------------------------------------------------------------------------
'* Win32_OnBoardDevice クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32OnBoardDevice = "Win32_OnBoardDevice"

'*-----------------------------------------------------------------------------
'* Win32_OperatingSystemAutochkSetting クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32OperatingSystemAutochkSetting = "Win32_OperatingSystemAutochkSetting"

'*-----------------------------------------------------------------------------
'* Win32_OperatingSystemQFE クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32OperatingSystemQFE = "Win32_OperatingSystemQFE"

'*-----------------------------------------------------------------------------
'* Win32_OptionalFeature クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32OptionalFeature = "Win32_OptionalFeature"

'*-----------------------------------------------------------------------------
'* Win32_OSRecoveryConfiguration クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32OSRecoveryConfiguration = "Win32_OSRecoveryConfiguration"

'*-----------------------------------------------------------------------------
'* Win32_PageFile クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PageFile = "Win32_PageFile"

'*-----------------------------------------------------------------------------
'* Win32_PageFileElementSetting クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PageFileElementSetting = "Win32_PageFileElementSetting"

'*-----------------------------------------------------------------------------
'* Win32_PageFileSetting クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PageFileSetting = "Win32_PageFileSetting"

'*-----------------------------------------------------------------------------
'* Win32_PageFileUsage クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PageFileUsage = "Win32_PageFileUsage"

'*-----------------------------------------------------------------------------
'* Win32_ParallelPort クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ParallelPort = "Win32_ParallelPort"

'*-----------------------------------------------------------------------------
'* Win32_Patch クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32Patch = "Win32_Patch"

'*-----------------------------------------------------------------------------
'* Win32_PatchFile クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PatchFile = "Win32_PatchFile"

'*-----------------------------------------------------------------------------
'* Win32_PatchPackage クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PatchPackage = "Win32_PatchPackage"

'*-----------------------------------------------------------------------------
'* Win32_PCMCIAController クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PCMCIAController = "Win32_PCMCIAController"

'*-----------------------------------------------------------------------------
'* Win32_PhysicalMedia クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PhysicalMedia = "Win32_PhysicalMedia"

'*-----------------------------------------------------------------------------
'* Win32_PhysicalMemory クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PhysicalMemory = "Win32_PhysicalMemory"

'*-----------------------------------------------------------------------------
'* Win32_PhysicalMemoryArray クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PhysicalMemoryArray = "Win32_PhysicalMemoryArray"

'*-----------------------------------------------------------------------------
'* Win32_PhysicalMemoryLocation クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PhysicalMemoryLocation = "Win32_PhysicalMemoryLocation"

'*-----------------------------------------------------------------------------
'* Win32_PNPAllocatedResource クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PNPAllocatedResource = "Win32_PNPAllocatedResource"

'*-----------------------------------------------------------------------------
'* Win32_PnPDevice クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PnPDevice = "Win32_PnPDevice"

'*-----------------------------------------------------------------------------
'* Win32_PnPSignedDriver クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PnPSignedDriver = "Win32_PnPSignedDriver"

'*-----------------------------------------------------------------------------
'* Win32_PnPSignedDriverCIMDataFile クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PnPSignedDriverCIMDataFile = "Win32_PnPSignedDriverCIMDataFile"

'*-----------------------------------------------------------------------------
'* Win32_PointingDevice クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PointingDevice = "Win32_PointingDevice"

'*-----------------------------------------------------------------------------
'* Win32_PortableBattery クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PortableBattery = "Win32_PortableBattery"

'*-----------------------------------------------------------------------------
'* Win32_PortConnector クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PortConnector = "Win32_PortConnector"

'*-----------------------------------------------------------------------------
'* Win32_PortResource クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PortResource = "Win32_PortResource"

'*-----------------------------------------------------------------------------
'* Win32_POTSModem クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32POTSModem = "Win32_POTSModem"

'*-----------------------------------------------------------------------------
'* Win32_POTSModemToSerialPort クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32POTSModemToSerialPort = "Win32_POTSModemToSerialPort"

'*-----------------------------------------------------------------------------
'* Win32_PowerManagementEvent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PowerManagementEvent = "Win32_PowerManagementEvent"

'*-----------------------------------------------------------------------------
'* Win32_PrinterConfiguration クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PrinterConfiguration = "Win32_PrinterConfiguration"

'*-----------------------------------------------------------------------------
'* Win32_PrinterController クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PrinterController = "Win32_PrinterController"

'*-----------------------------------------------------------------------------
'* Win32_PrinterDriver クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PrinterDriver = "Win32_PrinterDriver"

'*-----------------------------------------------------------------------------
'* Win32_PrinterDriverDll クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PrinterDriverDll = "Win32_PrinterDriverDll"

'*-----------------------------------------------------------------------------
'* Win32_PrinterSetting クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PrinterSetting = "Win32_PrinterSetting"

'*-----------------------------------------------------------------------------
'* Win32_PrinterShare クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PrinterShare = "Win32_PrinterShare"

'*-----------------------------------------------------------------------------
'* Win32_PrintJob クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PrintJob = "Win32_PrintJob"

'*-----------------------------------------------------------------------------
'* Win32_PrivilegesStatus クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PrivilegesStatus = "Win32_PrivilegesStatus"

'*-----------------------------------------------------------------------------
'* Win32_ProcessStartTrace クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ProcessStartTrace = "Win32_ProcessStartTrace"

'*-----------------------------------------------------------------------------
'* Win32_ProcessStartup クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ProcessStartup = "Win32_ProcessStartup"

'*-----------------------------------------------------------------------------
'* Win32_ProcessStopTrace クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ProcessStopTrace = "Win32_ProcessStopTrace"

'*-----------------------------------------------------------------------------
'* Win32_ProcessTrace クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ProcessTrace = "Win32_ProcessTrace"

'*-----------------------------------------------------------------------------
'* Win32_ProductCheck クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ProductCheck = "Win32_ProductCheck"

'*-----------------------------------------------------------------------------
'* Win32_ProductResource クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ProductResource = "Win32_ProductResource"

'*-----------------------------------------------------------------------------
'* Win32_ProductSoftwareFeatures クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ProductSoftwareFeatures = "Win32_ProductSoftwareFeatures"

'*-----------------------------------------------------------------------------
'* Win32_ProgIDSpecification クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ProgIDSpecification = "Win32_ProgIDSpecification"

'*-----------------------------------------------------------------------------
'* Win32_ProgramGroupContents クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ProgramGroupContents = "Win32_ProgramGroupContents"

'*-----------------------------------------------------------------------------
'* Win32_ProgramGroupOrItem クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ProgramGroupOrItem = "Win32_ProgramGroupOrItem"

'*-----------------------------------------------------------------------------
'* Win32_Property クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32Property = "Win32_Property"

'*-----------------------------------------------------------------------------
'* Win32_ProtocolBinding クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ProtocolBinding = "Win32_ProtocolBinding"

'*-----------------------------------------------------------------------------
'* Win32_PublishComponentAction クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PublishComponentAction = "Win32_PublishComponentAction"

'*-----------------------------------------------------------------------------
'* Win32_QuickFixEngineering クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32QuickFixEngineering = "Win32_QuickFixEngineering"

'*-----------------------------------------------------------------------------
'* Win32_QuotaSetting クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32QuotaSetting = "Win32_QuotaSetting"

'*-----------------------------------------------------------------------------
'* Win32_Refrigeration クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32Refrigeration = "Win32_Refrigeration"

'*-----------------------------------------------------------------------------
'* Win32_Registry クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32Registry = "Win32_Registry"

'*-----------------------------------------------------------------------------
'* Win32_RegistryAction クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32RegistryAction = "Win32_RegistryAction"

'*-----------------------------------------------------------------------------
'* Win32_Reliability クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32Reliability = "Win32_Reliability"

'*-----------------------------------------------------------------------------
'* Win32_ReliabilityRecords クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ReliabilityRecords = "Win32_ReliabilityRecords"

'*-----------------------------------------------------------------------------
'* Win32_ReliabilityStabilityMetrics クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ReliabilityStabilityMetrics = "Win32_ReliabilityStabilityMetrics"

'*-----------------------------------------------------------------------------
'* Win32_RemoveFileAction クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32RemoveFileAction = "Win32_RemoveFileAction"

'*-----------------------------------------------------------------------------
'* Win32_RemoveIniAction クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32RemoveIniAction = "Win32_RemoveIniAction"

'*-----------------------------------------------------------------------------
'* Win32_ReserveCost クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ReserveCost = "Win32_ReserveCost"

'*-----------------------------------------------------------------------------
'* Win32_RoamingProfileBackgroundUploadParams クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32RoamingProfileBackgroundUploadParams = "Win32_RoamingProfileBackgroundUploadParams"

'*-----------------------------------------------------------------------------
'* Win32_RoamingProfileMachineConfiguration クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32RoamingProfileMachineConfiguration = "Win32_RoamingProfileMachineConfiguration"

'*-----------------------------------------------------------------------------
'* Win32_RoamingProfileSlowLinkParams クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32RoamingProfileSlowLinkParams = "Win32_RoamingProfileSlowLinkParams"

'*-----------------------------------------------------------------------------
'* Win32_RoamingProfileUserConfiguration クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32RoamingProfileUserConfiguration = "Win32_RoamingProfileUserConfiguration"

'*-----------------------------------------------------------------------------
'* Win32_RoamingUserHealthConfiguration クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32RoamingUserHealthConfiguration = "Win32_RoamingUserHealthConfiguration"

'*-----------------------------------------------------------------------------
'* Win32_SCSIController クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SCSIController = "Win32_SCSIController"

'*-----------------------------------------------------------------------------
'* Win32_SCSIControllerDevice クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SCSIControllerDevice = "Win32_SCSIControllerDevice"

'*-----------------------------------------------------------------------------
'* Win32_SecurityDescriptor クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SecurityDescriptor = "Win32_SecurityDescriptor"

'*-----------------------------------------------------------------------------
'* Win32_SecurityDescriptorHelper クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SecurityDescriptorHelper = "Win32_SecurityDescriptorHelper"

'*-----------------------------------------------------------------------------
'* Win32_SecuritySetting クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SecuritySetting = "Win32_SecuritySetting"

'*-----------------------------------------------------------------------------
'* Win32_SecuritySettingAccess クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SecuritySettingAccess = "Win32_SecuritySettingAccess"

'*-----------------------------------------------------------------------------
'* Win32_SecuritySettingAuditing クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SecuritySettingAuditing = "Win32_SecuritySettingAuditing"

'*-----------------------------------------------------------------------------
'* Win32_SecuritySettingGroup クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SecuritySettingGroup = "Win32_SecuritySettingGroup"

'*-----------------------------------------------------------------------------
'* Win32_SecuritySettingOfLogicalFile クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SecuritySettingOfLogicalFile = "Win32_SecuritySettingOfLogicalFile"

'*-----------------------------------------------------------------------------
'* Win32_SecuritySettingOfLogicalShare クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SecuritySettingOfLogicalShare = "Win32_SecuritySettingOfLogicalShare"

'*-----------------------------------------------------------------------------
'* Win32_SecuritySettingOfObject クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SecuritySettingOfObject = "Win32_SecuritySettingOfObject"

'*-----------------------------------------------------------------------------
'* Win32_SecuritySettingOwner クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SecuritySettingOwner = "Win32_SecuritySettingOwner"

'*-----------------------------------------------------------------------------
'* Win32_SelfRegModuleAction クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SelfRegModuleAction = "Win32_SelfRegModuleAction"

'*-----------------------------------------------------------------------------
'* Win32_SerialPort クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SerialPort = "Win32_SerialPort"

'*-----------------------------------------------------------------------------
'* Win32_SerialPortConfiguration クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SerialPortConfiguration = "Win32_SerialPortConfiguration"

'*-----------------------------------------------------------------------------
'* Win32_SerialPortSetting クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SerialPortSetting = "Win32_SerialPortSetting"

'*-----------------------------------------------------------------------------
'* Win32_ServerConnection クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ServerConnection = "Win32_ServerConnection"

'*-----------------------------------------------------------------------------
'* Win32_ServerFeature クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ServerFeature = "Win32_ServerFeature"

'*-----------------------------------------------------------------------------
'* Win32_ServerSession クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ServerSession = "Win32_ServerSession"

'*-----------------------------------------------------------------------------
'* Win32_ServiceControl クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ServiceControl = "Win32_ServiceControl"

'*-----------------------------------------------------------------------------
'* Win32_ServiceSpecification クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ServiceSpecification = "Win32_ServiceSpecification"

'*-----------------------------------------------------------------------------
'* Win32_ServiceSpecificationService クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ServiceSpecificationService = "Win32_ServiceSpecificationService"

'*-----------------------------------------------------------------------------
'* Win32_Session クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32Session = "Win32_Session"

'*-----------------------------------------------------------------------------
'* Win32_SessionConnection クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SessionConnection = "Win32_SessionConnection"

'*-----------------------------------------------------------------------------
'* Win32_SessionProcess クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SessionProcess = "Win32_SessionProcess"

'*-----------------------------------------------------------------------------
'* Win32_SessionResource クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SessionResource = "Win32_SessionResource"

'*-----------------------------------------------------------------------------
'* Win32_SettingCheck クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SettingCheck = "Win32_SettingCheck"

'*-----------------------------------------------------------------------------
'* Win32_ShadowBy クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ShadowBy = "Win32_ShadowBy"

'*-----------------------------------------------------------------------------
'* Win32_ShadowContext クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ShadowContext = "Win32_ShadowContext"

'*-----------------------------------------------------------------------------
'* Win32_ShadowCopy クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ShadowCopy = "Win32_ShadowCopy"

'*-----------------------------------------------------------------------------
'* Win32_ShadowDiffVolumeSupport クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ShadowDiffVolumeSupport = "Win32_ShadowDiffVolumeSupport"

'*-----------------------------------------------------------------------------
'* Win32_ShadowFor クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ShadowFor = "Win32_ShadowFor"

'*-----------------------------------------------------------------------------
'* Win32_ShadowOn クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ShadowOn = "Win32_ShadowOn"

'*-----------------------------------------------------------------------------
'* Win32_ShadowProvider クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ShadowProvider = "Win32_ShadowProvider"

'*-----------------------------------------------------------------------------
'* Win32_ShadowStorage クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ShadowStorage = "Win32_ShadowStorage"

'*-----------------------------------------------------------------------------
'* Win32_ShadowVolumeSupport クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ShadowVolumeSupport = "Win32_ShadowVolumeSupport"

'*-----------------------------------------------------------------------------
'* Win32_Share クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32Share = "Win32_Share"

'*-----------------------------------------------------------------------------
'* Win32_ShareToDirectory クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ShareToDirectory = "Win32_ShareToDirectory"

'*-----------------------------------------------------------------------------
'* Win32_ShortcutAction クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ShortcutAction = "Win32_ShortcutAction"

'*-----------------------------------------------------------------------------
'* Win32_ShortcutFile クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ShortcutFile = "Win32_ShortcutFile"

'*-----------------------------------------------------------------------------
'* Win32_ShortcutSAP クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ShortcutSAP = "Win32_ShortcutSAP"

'*-----------------------------------------------------------------------------
'* Win32_SID クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SID = "Win32_SID"

'*-----------------------------------------------------------------------------
'* Win32_SIDandAttributes クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SIDandAttributes = "Win32_SIDandAttributes"

'*-----------------------------------------------------------------------------
'* Win32_SMBIOSMemory クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SMBIOSMemory = "Win32_SMBIOSMemory"

'*-----------------------------------------------------------------------------
'* Win32_SoftwareElement クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SoftwareElement = "Win32_SoftwareElement"

'*-----------------------------------------------------------------------------
'* Win32_SoftwareElementAction クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SoftwareElementAction = "Win32_SoftwareElementAction"

'*-----------------------------------------------------------------------------
'* Win32_SoftwareElementCheck クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SoftwareElementCheck = "Win32_SoftwareElementCheck"

'*-----------------------------------------------------------------------------
'* Win32_SoftwareElementCondition クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SoftwareElementCondition = "Win32_SoftwareElementCondition"

'*-----------------------------------------------------------------------------
'* Win32_SoftwareElementResource クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SoftwareElementResource = "Win32_SoftwareElementResource"

'*-----------------------------------------------------------------------------
'* Win32_SoftwareFeature クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SoftwareFeature = "Win32_SoftwareFeature"

'*-----------------------------------------------------------------------------
'* Win32_SoftwareFeatureAction クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SoftwareFeatureAction = "Win32_SoftwareFeatureAction"

'*-----------------------------------------------------------------------------
'* Win32_SoftwareFeatureCheck クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SoftwareFeatureCheck = "Win32_SoftwareFeatureCheck"

'*-----------------------------------------------------------------------------
'* Win32_SoftwareFeatureParent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SoftwareFeatureParent = "Win32_SoftwareFeatureParent"

'*-----------------------------------------------------------------------------
'* Win32_SoftwareFeatureSoftwareElements クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SoftwareFeatureSoftwareElements = "Win32_SoftwareFeatureSoftwareElements"

'*-----------------------------------------------------------------------------
'* Win32_SoundDevice クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SoundDevice = "Win32_SoundDevice"

'*-----------------------------------------------------------------------------
'* Win32_StartupCommand クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32StartupCommand = "Win32_StartupCommand"

'*-----------------------------------------------------------------------------
'* Win32_SubDirectory クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SubDirectory = "Win32_SubDirectory"

'*-----------------------------------------------------------------------------
'* Win32_SubSession クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SubSession = "Win32_SubSession"

'*-----------------------------------------------------------------------------
'* Win32_SystemAccount クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SystemAccount = "Win32_SystemAccount"

'*-----------------------------------------------------------------------------
'* Win32_SystemBIOS クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SystemBIOS = "Win32_SystemBIOS"

'*-----------------------------------------------------------------------------
'* Win32_SystemBootConfiguration クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SystemBootConfiguration = "Win32_SystemBootConfiguration"

'*-----------------------------------------------------------------------------
'* Win32_SystemConfigurationChangeEvent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SystemConfigurationChangeEvent = "Win32_SystemConfigurationChangeEvent"

'*-----------------------------------------------------------------------------
'* Win32_SystemDesktop クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SystemDesktop = "Win32_SystemDesktop"

'*-----------------------------------------------------------------------------
'* Win32_SystemDevices クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SystemDevices = "Win32_SystemDevices"

'*-----------------------------------------------------------------------------
'* Win32_SystemDriver クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SystemDriver = "Win32_SystemDriver"

'*-----------------------------------------------------------------------------
'* Win32_SystemDriverPNPEntity クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SystemDriverPNPEntity = "Win32_SystemDriverPNPEntity"

'*-----------------------------------------------------------------------------
'* Win32_SystemEnclosure クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SystemEnclosure = "Win32_SystemEnclosure"

'*-----------------------------------------------------------------------------
'* Win32_SystemLoadOrderGroups クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SystemLoadOrderGroups = "Win32_SystemLoadOrderGroups"

'*-----------------------------------------------------------------------------
'* Win32_SystemMemoryResource クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SystemMemoryResource = "Win32_SystemMemoryResource"

'*-----------------------------------------------------------------------------
'* Win32_SystemNetworkConnections クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SystemNetworkConnections = "Win32_SystemNetworkConnections"

'*-----------------------------------------------------------------------------
'* Win32_SystemOperatingSystem クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SystemOperatingSystem = "Win32_SystemOperatingSystem"

'*-----------------------------------------------------------------------------
'* Win32_SystemPartitions クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SystemPartitions = "Win32_SystemPartitions"

'*-----------------------------------------------------------------------------
'* Win32_SystemProcesses クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SystemProcesses = "Win32_SystemProcesses"

'*-----------------------------------------------------------------------------
'* Win32_SystemProgramGroups クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SystemProgramGroups = "Win32_SystemProgramGroups"

'*-----------------------------------------------------------------------------
'* Win32_SystemResources クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SystemResources = "Win32_SystemResources"

'*-----------------------------------------------------------------------------
'* Win32_SystemServices クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SystemServices = "Win32_SystemServices"

'*-----------------------------------------------------------------------------
'* Win32_SystemSetting クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SystemSetting = "Win32_SystemSetting"

'*-----------------------------------------------------------------------------
'* Win32_SystemSlot クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SystemSlot = "Win32_SystemSlot"

'*-----------------------------------------------------------------------------
'* Win32_SystemSystemDriver クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SystemSystemDriver = "Win32_SystemSystemDriver"

'*-----------------------------------------------------------------------------
'* Win32_SystemTimeZone クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SystemTimeZone = "Win32_SystemTimeZone"

'*-----------------------------------------------------------------------------
'* Win32_SystemTrace クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SystemTrace = "Win32_SystemTrace"

'*-----------------------------------------------------------------------------
'* Win32_SystemUsers クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32SystemUsers = "Win32_SystemUsers"

'*-----------------------------------------------------------------------------
'* Win32_TapeDrive クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32TapeDrive = "Win32_TapeDrive"

'*-----------------------------------------------------------------------------
'* Win32_TCPIPPrinterPort クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32TCPIPPrinterPort = "Win32_TCPIPPrinterPort"

'*-----------------------------------------------------------------------------
'* Win32_TemperatureProbe クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32TemperatureProbe = "Win32_TemperatureProbe"

'*-----------------------------------------------------------------------------
'* Win32_TerminalService クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32TerminalService = "Win32_TerminalService"

'*-----------------------------------------------------------------------------
'* Win32_Thread クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32Thread = "Win32_Thread"

'*-----------------------------------------------------------------------------
'* Win32_ThreadStartTrace クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ThreadStartTrace = "Win32_ThreadStartTrace"

'*-----------------------------------------------------------------------------
'* Win32_ThreadStopTrace クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ThreadStopTrace = "Win32_ThreadStopTrace"

'*-----------------------------------------------------------------------------
'* Win32_ThreadTrace クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32ThreadTrace = "Win32_ThreadTrace"

'*-----------------------------------------------------------------------------
'* Win32_TimeZone クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32TimeZone = "Win32_TimeZone"

'*-----------------------------------------------------------------------------
'* Win32_TokenGroups クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32TokenGroups = "Win32_TokenGroups"

'*-----------------------------------------------------------------------------
'* Win32_TokenPrivileges クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32TokenPrivileges = "Win32_TokenPrivileges"

'*-----------------------------------------------------------------------------
'* Win32_Trustee クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32Trustee = "Win32_Trustee"

'*-----------------------------------------------------------------------------
'* Win32_TypeLibraryAction クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32TypeLibraryAction = "Win32_TypeLibraryAction"

'*-----------------------------------------------------------------------------
'* Win32_USBController クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32USBController = "Win32_USBController"

'*-----------------------------------------------------------------------------
'* Win32_USBControllerDevice クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32USBControllerDevice = "Win32_USBControllerDevice"

'*-----------------------------------------------------------------------------
'* Win32_USBHub クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32USBHub = "Win32_USBHub"

'*-----------------------------------------------------------------------------
'* Win32_UserDesktop クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32UserDesktop = "Win32_UserDesktop"

'*-----------------------------------------------------------------------------
'* Win32_UserInDomain クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32UserInDomain = "Win32_UserInDomain"

'*-----------------------------------------------------------------------------
'* Win32_UserProfile クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32UserProfile = "Win32_UserProfile"

'*-----------------------------------------------------------------------------
'* Win32_UserStateConfigurationControls クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32UserStateConfigurationControls = "Win32_UserStateConfigurationControls"

'*-----------------------------------------------------------------------------
'* Win32_UTCTime クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32UTCTime = "Win32_UTCTime"

'*-----------------------------------------------------------------------------
'* Win32_VideoConfiguration クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32VideoConfiguration = "Win32_VideoConfiguration"

'*-----------------------------------------------------------------------------
'* Win32_VideoController クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32VideoController = "Win32_VideoController"

'*-----------------------------------------------------------------------------
'* Win32_VideoSettings クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32VideoSettings = "Win32_VideoSettings"

'*-----------------------------------------------------------------------------
'* Win32_VoltageProbe クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32VoltageProbe = "Win32_VoltageProbe"

'*-----------------------------------------------------------------------------
'* Win32_Volume クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32Volume = "Win32_Volume"

'*-----------------------------------------------------------------------------
'* Win32_VolumeChangeEvent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32VolumeChangeEvent = "Win32_VolumeChangeEvent"

'*-----------------------------------------------------------------------------
'* Win32_VolumeQuota クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32VolumeQuota = "Win32_VolumeQuota"

'*-----------------------------------------------------------------------------
'* Win32_VolumeQuotaSetting クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32VolumeQuotaSetting = "Win32_VolumeQuotaSetting"

'*-----------------------------------------------------------------------------
'* Win32_VolumeUserQuota クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32VolumeUserQuota = "Win32_VolumeUserQuota"

'*-----------------------------------------------------------------------------
'* Win32_WinSAT クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32WinSAT = "Win32_WinSAT"

'*-----------------------------------------------------------------------------
'* Win32_WMIElementSetting クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32WMIElementSetting = "Win32_WMIElementSetting"

'*-----------------------------------------------------------------------------
'* Win32_WMISetting クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32WMISetting = "Win32_WMISetting"

'*-----------------------------------------------------------------------------
'* Win32_Perf クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32Perf = "Win32_Perf"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedData = "Win32_PerfFormattedData"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_AFDCounters_MicrosoftWinsockBSP クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataAFDCountersMicrosoftWinsockBSP = "Win32_PerfFormattedData_AFDCounters_MicrosoftWinsockBSP"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_APPPOOLCountersProvider_APPPOOLWAS クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataAPPPOOLCountersProviderAPPPOOLWAS = "Win32_PerfFormattedData_APPPOOLCountersProvider_APPPOOLWAS"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_ASP_ActiveServerPages クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataASPActiveServerPages = "Win32_PerfFormattedData_ASP_ActiveServerPages"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_ASPNET_ASPNET クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataASPNETASPNET = "Win32_PerfFormattedData_ASPNET_ASPNET"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_ASPNET_ASPNETApplications クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataASPNETASPNETApplications = "Win32_PerfFormattedData_ASPNET_ASPNETApplications"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_ASPNET2050727_ASPNETAppsv2050727 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataASPNET2050727ASPNETAppsv2050727 = "Win32_PerfFormattedData_ASPNET2050727_ASPNETAppsv2050727"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_ASPNET2050727_ASPNETv2050727 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataASPNET2050727ASPNETv2050727 = "Win32_PerfFormattedData_ASPNET2050727_ASPNETv2050727"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_ASPNET4030319_ASPNETAppsv4030319 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataASPNET4030319ASPNETAppsv4030319 = "Win32_PerfFormattedData_ASPNET4030319_ASPNETAppsv4030319"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_ASPNET4030319_ASPNETv4030319 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataASPNET4030319ASPNETv4030319 = "Win32_PerfFormattedData_ASPNET4030319_ASPNETv4030319"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_aspnetstate_ASPNETStateService クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataaspnetstateASPNETStateService = "Win32_PerfFormattedData_aspnetstate_ASPNETStateService"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_AuthorizationManager_AuthorizationManagerApplications クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataAuthorizationManagerAuthorizationManagerApplications = "Win32_PerfFormattedData_AuthorizationManager_AuthorizationManagerApplications"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_BalancerStats_HyperVDynamicMemoryBalancer クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataBalancerStatsHyperVDynamicMemoryBalancer = "Win32_PerfFormattedData_BalancerStats_HyperVDynamicMemoryBalancer"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_BalancerStats_HyperVDynamicMemoryVM クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataBalancerStatsHyperVDynamicMemoryVM = "Win32_PerfFormattedData_BalancerStats_HyperVDynamicMemoryVM"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_BITS_BITSNetUtilization クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataBITSBITSNetUtilization = "Win32_PerfFormattedData_BITS_BITSNetUtilization"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_DNS64Global クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersDNS64Global = "Win32_PerfFormattedData_Counters_DNS64Global"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_EventTracingforWindows クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersEventTracingforWindows = "Win32_PerfFormattedData_Counters_EventTracingforWindows"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_EventTracingforWindowsSession クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersEventTracingforWindowsSession = "Win32_PerfFormattedData_Counters_EventTracingforWindowsSession"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_FileSystemDiskActivity クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersFileSystemDiskActivity = "Win32_PerfFormattedData_Counters_FileSystemDiskActivity"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_GenericIKEv1AuthIPandIKEv2 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersGenericIKEv1AuthIPandIKEv2 = "Win32_PerfFormattedData_Counters_GenericIKEv1AuthIPandIKEv2"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_HTTPService クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersHTTPService = "Win32_PerfFormattedData_Counters_HTTPService"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_HTTPServiceRequestQueues クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersHTTPServiceRequestQueues = "Win32_PerfFormattedData_Counters_HTTPServiceRequestQueues"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_HTTPServiceUrlGroups クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersHTTPServiceUrlGroups = "Win32_PerfFormattedData_Counters_HTTPServiceUrlGroups"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_HyperVDynamicMemoryIntegrationService クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersHyperVDynamicMemoryIntegrationService = "Win32_PerfFormattedData_Counters_HyperVDynamicMemoryIntegrationService"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_HyperVVirtualMachineBusPipes クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersHyperVVirtualMachineBusPipes = "Win32_PerfFormattedData_Counters_HyperVVirtualMachineBusPipes"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_IPHTTPSGlobal クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersIPHTTPSGlobal = "Win32_PerfFormattedData_Counters_IPHTTPSGlobal"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_IPHTTPSSession クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersIPHTTPSSession = "Win32_PerfFormattedData_Counters_IPHTTPSSession"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_IPsecAuthIPIPv4 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersIPsecAuthIPIPv4 = "Win32_PerfFormattedData_Counters_IPsecAuthIPIPv4"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_IPsecAuthIPIPv6 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersIPsecAuthIPIPv6 = "Win32_PerfFormattedData_Counters_IPsecAuthIPIPv6"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_IPsecConnections クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersIPsecConnections = "Win32_PerfFormattedData_Counters_IPsecConnections"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_IPsecDoSProtection クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersIPsecDoSProtection = "Win32_PerfFormattedData_Counters_IPsecDoSProtection"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_IPsecDriver クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersIPsecDriver = "Win32_PerfFormattedData_Counters_IPsecDriver"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_IPsecIKEv1IPv4 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersIPsecIKEv1IPv4 = "Win32_PerfFormattedData_Counters_IPsecIKEv1IPv4"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_IPsecIKEv1IPv6 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersIPsecIKEv1IPv6 = "Win32_PerfFormattedData_Counters_IPsecIKEv1IPv6"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_IPsecIKEv2IPv4 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersIPsecIKEv2IPv4 = "Win32_PerfFormattedData_Counters_IPsecIKEv2IPv4"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_IPsecIKEv2IPv6 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersIPsecIKEv2IPv6 = "Win32_PerfFormattedData_Counters_IPsecIKEv2IPv6"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_Netlogon クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersNetlogon = "Win32_PerfFormattedData_Counters_Netlogon"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_NetworkQoSPolicy クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersNetworkQoSPolicy = "Win32_PerfFormattedData_Counters_NetworkQoSPolicy"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_PacerFlow クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersPacerFlow = "Win32_PerfFormattedData_Counters_PacerFlow"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_PacerPipe クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersPacerPipe = "Win32_PerfFormattedData_Counters_PacerPipe"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_PacketDirectECUtilization クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersPacketDirectECUtilization = "Win32_PerfFormattedData_Counters_PacketDirectECUtilization"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_PacketDirectQueueDepth クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersPacketDirectQueueDepth = "Win32_PerfFormattedData_Counters_PacketDirectQueueDepth"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_PacketDirectReceiveCounters クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersPacketDirectReceiveCounters = "Win32_PerfFormattedData_Counters_PacketDirectReceiveCounters"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_PacketDirectReceiveFilters クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersPacketDirectReceiveFilters = "Win32_PerfFormattedData_Counters_PacketDirectReceiveFilters"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_PacketDirectTransmitCounters クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersPacketDirectTransmitCounters = "Win32_PerfFormattedData_Counters_PacketDirectTransmitCounters"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_PerProcessorNetworkActivityCycles クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersPerProcessorNetworkActivityCycles = "Win32_PerfFormattedData_Counters_PerProcessorNetworkActivityCycles"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_PerProcessorNetworkInterfaceCardActivity クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersPerProcessorNetworkInterfaceCardActivity = "Win32_PerfFormattedData_Counters_PerProcessorNetworkInterfaceCardActivity"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_PhysicalNetworkInterfaceCardActivity クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersPhysicalNetworkInterfaceCardActivity = "Win32_PerfFormattedData_Counters_PhysicalNetworkInterfaceCardActivity"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_PowerShellWorkflow クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersPowerShellWorkflow = "Win32_PerfFormattedData_Counters_PowerShellWorkflow"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_ProcessorInformation クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersProcessorInformation = "Win32_PerfFormattedData_Counters_ProcessorInformation"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_RDMAActivity クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersRDMAActivity = "Win32_PerfFormattedData_Counters_RDMAActivity"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_RemoteFXGraphics クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersRemoteFXGraphics = "Win32_PerfFormattedData_Counters_RemoteFXGraphics"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_RemoteFXNetwork クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersRemoteFXNetwork = "Win32_PerfFormattedData_Counters_RemoteFXNetwork"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_SMBClientShares クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersSMBClientShares = "Win32_PerfFormattedData_Counters_SMBClientShares"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_SMBServer クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersSMBServer = "Win32_PerfFormattedData_Counters_SMBServer"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_SMBServerSessions クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersSMBServerSessions = "Win32_PerfFormattedData_Counters_SMBServerSessions"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_SMBServerShares クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersSMBServerShares = "Win32_PerfFormattedData_Counters_SMBServerShares"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_StorageSpacesTier クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersStorageSpacesTier = "Win32_PerfFormattedData_Counters_StorageSpacesTier"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_StorageSpacesWriteCache クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersStorageSpacesWriteCache = "Win32_PerfFormattedData_Counters_StorageSpacesWriteCache"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_Synchronization クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersSynchronization = "Win32_PerfFormattedData_Counters_Synchronization"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_SynchronizationNuma クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersSynchronizationNuma = "Win32_PerfFormattedData_Counters_SynchronizationNuma"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_TeredoClient クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersTeredoClient = "Win32_PerfFormattedData_Counters_TeredoClient"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_TeredoRelay クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersTeredoRelay = "Win32_PerfFormattedData_Counters_TeredoRelay"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_TeredoServer クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersTeredoServer = "Win32_PerfFormattedData_Counters_TeredoServer"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_ThermalZoneInformation クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersThermalZoneInformation = "Win32_PerfFormattedData_Counters_ThermalZoneInformation"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_WFP クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersWFP = "Win32_PerfFormattedData_Counters_WFP"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_WFPv4 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersWFPv4 = "Win32_PerfFormattedData_Counters_WFPv4"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_WFPv6 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersWFPv6 = "Win32_PerfFormattedData_Counters_WFPv6"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_WSManQuotaStatistics クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersWSManQuotaStatistics = "Win32_PerfFormattedData_Counters_WSManQuotaStatistics"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_XHCICommonBuffer クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersXHCICommonBuffer = "Win32_PerfFormattedData_Counters_XHCICommonBuffer"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_XHCIInterrupter クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersXHCIInterrupter = "Win32_PerfFormattedData_Counters_XHCIInterrupter"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Counters_XHCITransferRing クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataCountersXHCITransferRing = "Win32_PerfFormattedData_Counters_XHCITransferRing"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_DdmCounterProvider_RAS クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataDdmCounterProviderRAS = "Win32_PerfFormattedData_DdmCounterProvider_RAS"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_DeliveryOptimization_DeliveryOptimizationSwarm クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataDeliveryOptimizationDeliveryOptimizationSwarm = "Win32_PerfFormattedData_DeliveryOptimization_DeliveryOptimizationSwarm"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_DistributedRoutingTablePerf_DistributedRoutingTable クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataDistributedRoutingTablePerfDistributedRoutingTable = "Win32_PerfFormattedData_DistributedRoutingTablePerf_DistributedRoutingTable"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_ESENT_Database クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataESENTDatabase = "Win32_PerfFormattedData_ESENT_Database"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_ESENT_DatabaseInstances クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataESENTDatabaseInstances = "Win32_PerfFormattedData_ESENT_DatabaseInstances"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_ESENT_DatabaseTableClasses クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataESENTDatabaseTableClasses = "Win32_PerfFormattedData_ESENT_DatabaseTableClasses"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_EthernetPerfProvider_HyperVLegacyNetworkAdapter クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataEthernetPerfProviderHyperVLegacyNetworkAdapter = "Win32_PerfFormattedData_EthernetPerfProvider_HyperVLegacyNetworkAdapter"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_FaxService_FaxService クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataFaxServiceFaxService = "Win32_PerfFormattedData_FaxService_FaxService"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_ftpsvc_MicrosoftFTPService クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataftpsvcMicrosoftFTPService = "Win32_PerfFormattedData_ftpsvc_MicrosoftFTPService"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_GmoPerfProvider_HyperVVMSaveSnapshotandRestore クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataGmoPerfProviderHyperVVMSaveSnapshotandRestore = "Win32_PerfFormattedData_GmoPerfProvider_HyperVVMSaveSnapshotandRestore"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_HvStats_HyperVHypervisor クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataHvStatsHyperVHypervisor = "Win32_PerfFormattedData_HvStats_HyperVHypervisor"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_HvStats_HyperVHypervisorLogicalProcessor クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataHvStatsHyperVHypervisorLogicalProcessor = "Win32_PerfFormattedData_HvStats_HyperVHypervisorLogicalProcessor"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_HvStats_HyperVHypervisorPartition クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataHvStatsHyperVHypervisorPartition = "Win32_PerfFormattedData_HvStats_HyperVHypervisorPartition"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_HvStats_HyperVHypervisorRootPartition クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataHvStatsHyperVHypervisorRootPartition = "Win32_PerfFormattedData_HvStats_HyperVHypervisorRootPartition"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_HvStats_HyperVHypervisorRootVirtualProcessor クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataHvStatsHyperVHypervisorRootVirtualProcessor = "Win32_PerfFormattedData_HvStats_HyperVHypervisorRootVirtualProcessor"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_HvStats_HyperVHypervisorVirtualProcessor クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataHvStatsHyperVHypervisorVirtualProcessor = "Win32_PerfFormattedData_HvStats_HyperVHypervisorVirtualProcessor"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_IdePerfProvider_HyperVVirtualIDEController クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataIdePerfProviderHyperVVirtualIDEController = "Win32_PerfFormattedData_IdePerfProvider_HyperVVirtualIDEController"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_LocalSessionManager_TerminalServices クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataLocalSessionManagerTerminalServices = "Win32_PerfFormattedData_LocalSessionManager_TerminalServices"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Lsa_SecurityPerProcessStatistics クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataLsaSecurityPerProcessStatistics = "Win32_PerfFormattedData_Lsa_SecurityPerProcessStatistics"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Lsa_SecuritySystemWideStatistics クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataLsaSecuritySystemWideStatistics = "Win32_PerfFormattedData_Lsa_SecuritySystemWideStatistics"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_MicrosoftWindowsBitLockerDriverCountersProvider_BitLocker クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataMicrosoftWindowsBitLockerDriverCountersProviderBitLocker = "Win32_PerfFormattedData_MicrosoftWindowsBitLockerDriverCountersProvider_BitLocker"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_MicrosoftWindowsRemoteDesktopServicesRemoteFXSynth3dvsc_RemoteFXSynth3DVSCVMDevice クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataMicrosoftWindowsRemoteDesktopServicesRemoteFXSynth3dvscRemoteFXSynth3DVSCVMDevice = "Win32_PerfFormattedData_MicrosoftWindowsRemoteDesktopServicesRemoteFXSynth3dvsc_RemoteFXSynth3DVSCVMDevice"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_MicrosoftWindowsRemoteDesktopServicesRemoteFXSynth3dvsc_RemoteFXSynth3DVSCVMTransportChannel クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataMicrosoftWindowsRemoteDesktopServicesRemoteFXSynth3dvscRemoteFXSynth3DVSCVMTransportChannel = "Win32_PerfFormattedData_MicrosoftWindowsRemoteDesktopServicesRemoteFXSynth3dvsc_RemoteFXSynth3DVSCVMTransportChannel"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_MSDTC_DistributedTransactionCoordinator クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataMSDTCDistributedTransactionCoordinator = "Win32_PerfFormattedData_MSDTC_DistributedTransactionCoordinator"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_MSDTCBridge3000_MSDTCBridge3000 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataMSDTCBridge3000MSDTCBridge3000 = "Win32_PerfFormattedData_MSDTCBridge3000_MSDTCBridge3000"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_MSDTCBridge4000_MSDTCBridge4000 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataMSDTCBridge4000MSDTCBridge4000 = "Win32_PerfFormattedData_MSDTCBridge4000_MSDTCBridge4000"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_NETCLRData_NETCLRData クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataNETCLRDataNETCLRData = "Win32_PerfFormattedData_NETCLRData_NETCLRData"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_NETCLRNetworking_NETCLRNetworking クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataNETCLRNetworkingNETCLRNetworking = "Win32_PerfFormattedData_NETCLRNetworking_NETCLRNetworking"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_NETCLRNetworking4000_NETCLRNetworking4000 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataNETCLRNetworking4000NETCLRNetworking4000 = "Win32_PerfFormattedData_NETCLRNetworking4000_NETCLRNetworking4000"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_NETDataProviderforOracle_NETDataProviderforOracle クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataNETDataProviderforOracleNETDataProviderforOracle = "Win32_PerfFormattedData_NETDataProviderforOracle_NETDataProviderforOracle"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_NETDataProviderforSqlServer_NETDataProviderforSqlServer クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataNETDataProviderforSqlServerNETDataProviderforSqlServer = "Win32_PerfFormattedData_NETDataProviderforSqlServer_NETDataProviderforSqlServer"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_NETFramework_NETCLRExceptions クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataNETFrameworkNETCLRExceptions = "Win32_PerfFormattedData_NETFramework_NETCLRExceptions"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_NETFramework_NETCLRInterop クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataNETFrameworkNETCLRInterop = "Win32_PerfFormattedData_NETFramework_NETCLRInterop"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_NETFramework_NETCLRJit クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataNETFrameworkNETCLRJit = "Win32_PerfFormattedData_NETFramework_NETCLRJit"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_NETFramework_NETCLRLoading クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataNETFrameworkNETCLRLoading = "Win32_PerfFormattedData_NETFramework_NETCLRLoading"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_NETFramework_NETCLRLocksAndThreads クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataNETFrameworkNETCLRLocksAndThreads = "Win32_PerfFormattedData_NETFramework_NETCLRLocksAndThreads"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_NETFramework_NETCLRMemory クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataNETFrameworkNETCLRMemory = "Win32_PerfFormattedData_NETFramework_NETCLRMemory"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_NETFramework_NETCLRRemoting クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataNETFrameworkNETCLRRemoting = "Win32_PerfFormattedData_NETFramework_NETCLRRemoting"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_NETFramework_NETCLRSecurity クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataNETFrameworkNETCLRSecurity = "Win32_PerfFormattedData_NETFramework_NETCLRSecurity"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_NETMemoryCache40_NETMemoryCache40 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataNETMemoryCache40NETMemoryCache40 = "Win32_PerfFormattedData_NETMemoryCache40_NETMemoryCache40"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_NvspNicStats_HyperVVirtualNetworkAdapter クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataNvspNicStatsHyperVVirtualNetworkAdapter = "Win32_PerfFormattedData_NvspNicStats_HyperVVirtualNetworkAdapter"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_NvspPortStats_HyperVVirtualSwitchPort クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataNvspPortStatsHyperVVirtualSwitchPort = "Win32_PerfFormattedData_NvspPortStats_HyperVVirtualSwitchPort"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_NvspSwitchStats_HyperVVirtualSwitch クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataNvspSwitchStatsHyperVVirtualSwitch = "Win32_PerfFormattedData_NvspSwitchStats_HyperVVirtualSwitch"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_OfflineFiles_ClientSideCaching クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataOfflineFilesClientSideCaching = "Win32_PerfFormattedData_OfflineFiles_ClientSideCaching"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_OfflineFiles_OfflineFiles クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataOfflineFilesOfflineFiles = "Win32_PerfFormattedData_OfflineFiles_OfflineFiles"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_PeerDistSvc_BranchCache クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataPeerDistSvcBranchCache = "Win32_PerfFormattedData_PeerDistSvc_BranchCache"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_PeerNameResolutionProtocolPerf_PeerNameResolutionProtocol クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataPeerNameResolutionProtocolPerfPeerNameResolutionProtocol = "Win32_PerfFormattedData_PeerNameResolutionProtocolPerf_PeerNameResolutionProtocol"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_PerfDisk_LogicalDisk クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataPerfDiskLogicalDisk = "Win32_PerfFormattedData_PerfDisk_LogicalDisk"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_PerfDisk_PhysicalDisk クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataPerfDiskPhysicalDisk = "Win32_PerfFormattedData_PerfDisk_PhysicalDisk"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_PerfNet_Browser クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataPerfNetBrowser = "Win32_PerfFormattedData_PerfNet_Browser"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_PerfNet_Redirector クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataPerfNetRedirector = "Win32_PerfFormattedData_PerfNet_Redirector"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_PerfNet_Server クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataPerfNetServer = "Win32_PerfFormattedData_PerfNet_Server"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_PerfNet_ServerWorkQueues クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataPerfNetServerWorkQueues = "Win32_PerfFormattedData_PerfNet_ServerWorkQueues"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_PerfOS_Cache クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataPerfOSCache = "Win32_PerfFormattedData_PerfOS_Cache"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_PerfOS_Memory クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataPerfOSMemory = "Win32_PerfFormattedData_PerfOS_Memory"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_PerfOS_NUMANodeMemory クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataPerfOSNUMANodeMemory = "Win32_PerfFormattedData_PerfOS_NUMANodeMemory"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_PerfOS_Objects クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataPerfOSObjects = "Win32_PerfFormattedData_PerfOS_Objects"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_PerfOS_PagingFile クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataPerfOSPagingFile = "Win32_PerfFormattedData_PerfOS_PagingFile"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_PerfOS_Processor クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataPerfOSProcessor = "Win32_PerfFormattedData_PerfOS_Processor"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_PerfOS_System クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataPerfOSSystem = "Win32_PerfFormattedData_PerfOS_System"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_PerfProc_FullImage_Costly クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataPerfProcFullImageCostly = "Win32_PerfFormattedData_PerfProc_FullImage_Costly"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_PerfProc_Image_Costly クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataPerfProcImageCostly = "Win32_PerfFormattedData_PerfProc_Image_Costly"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_PerfProc_JobObject クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataPerfProcJobObject = "Win32_PerfFormattedData_PerfProc_JobObject"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_PerfProc_JobObjectDetails クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataPerfProcJobObjectDetails = "Win32_PerfFormattedData_PerfProc_JobObjectDetails"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_PerfProc_Process クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataPerfProcProcess = "Win32_PerfFormattedData_PerfProc_Process"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_PerfProc_ProcessAddressSpace_Costly クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataPerfProcProcessAddressSpaceCostly = "Win32_PerfFormattedData_PerfProc_ProcessAddressSpace_Costly"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_PerfProc_Thread クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataPerfProcThread = "Win32_PerfFormattedData_PerfProc_Thread"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_PerfProc_ThreadDetails_Costly クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataPerfProcThreadDetailsCostly = "Win32_PerfFormattedData_PerfProc_ThreadDetails_Costly"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_PowerMeterCounter_EnergyMeter クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataPowerMeterCounterEnergyMeter = "Win32_PerfFormattedData_PowerMeterCounter_EnergyMeter"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_PowerMeterCounter_PowerMeter クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataPowerMeterCounterPowerMeter = "Win32_PerfFormattedData_PowerMeterCounter_PowerMeter"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_rdyboost_ReadyBoostCache クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDatardyboostReadyBoostCache = "Win32_PerfFormattedData_rdyboost_ReadyBoostCache"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_RemoteAccess_RASPort クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataRemoteAccessRASPort = "Win32_PerfFormattedData_RemoteAccess_RASPort"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_RemoteAccess_RASTotal クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataRemoteAccessRASTotal = "Win32_PerfFormattedData_RemoteAccess_RASTotal"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_RemotePerfProvider_HyperVVMRemoting クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataRemotePerfProviderHyperVVMRemoting = "Win32_PerfFormattedData_RemotePerfProvider_HyperVVMRemoting"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_ServiceModel4000_ServiceModelEndpoint4000 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataServiceModel4000ServiceModelEndpoint4000 = "Win32_PerfFormattedData_ServiceModel4000_ServiceModelEndpoint4000"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_ServiceModel4000_ServiceModelOperation4000 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataServiceModel4000ServiceModelOperation4000 = "Win32_PerfFormattedData_ServiceModel4000_ServiceModelOperation4000"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_ServiceModel4000_ServiceModelService4000 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataServiceModel4000ServiceModelService4000 = "Win32_PerfFormattedData_ServiceModel4000_ServiceModelService4000"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_ServiceModelEndpoint3000_ServiceModelEndpoint3000 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataServiceModelEndpoint3000ServiceModelEndpoint3000 = "Win32_PerfFormattedData_ServiceModelEndpoint3000_ServiceModelEndpoint3000"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_ServiceModelOperation3000_ServiceModelOperation3000 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataServiceModelOperation3000ServiceModelOperation3000 = "Win32_PerfFormattedData_ServiceModelOperation3000_ServiceModelOperation3000"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_ServiceModelService3000_ServiceModelService3000 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataServiceModelService3000ServiceModelService3000 = "Win32_PerfFormattedData_ServiceModelService3000_ServiceModelService3000"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_SMSvcHost3000_SMSvcHost3000 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataSMSvcHost3000SMSvcHost3000 = "Win32_PerfFormattedData_SMSvcHost3000_SMSvcHost3000"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_SMSvcHost4000_SMSvcHost4000 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataSMSvcHost4000SMSvcHost4000 = "Win32_PerfFormattedData_SMSvcHost4000_SMSvcHost4000"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Spooler_PrintQueue クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataSpoolerPrintQueue = "Win32_PerfFormattedData_Spooler_PrintQueue"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_StorageStats_HyperVVirtualStorageDevice クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataStorageStatsHyperVVirtualStorageDevice = "Win32_PerfFormattedData_StorageStats_HyperVVirtualStorageDevice"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_TapiSrv_Telephony クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataTapiSrvTelephony = "Win32_PerfFormattedData_TapiSrv_Telephony"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_TBS_TBScounters クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataTBSTBScounters = "Win32_PerfFormattedData_TBS_TBScounters"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Tcpip_ICMP クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataTcpipICMP = "Win32_PerfFormattedData_Tcpip_ICMP"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Tcpip_ICMPv6 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataTcpipICMPv6 = "Win32_PerfFormattedData_Tcpip_ICMPv6"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Tcpip_IPv4 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataTcpipIPv4 = "Win32_PerfFormattedData_Tcpip_IPv4"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Tcpip_IPv6 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataTcpipIPv6 = "Win32_PerfFormattedData_Tcpip_IPv6"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Tcpip_NBTConnection クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataTcpipNBTConnection = "Win32_PerfFormattedData_Tcpip_NBTConnection"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Tcpip_NetworkAdapter クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataTcpipNetworkAdapter = "Win32_PerfFormattedData_Tcpip_NetworkAdapter"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Tcpip_NetworkInterface クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataTcpipNetworkInterface = "Win32_PerfFormattedData_Tcpip_NetworkInterface"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Tcpip_TCPv4 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataTcpipTCPv4 = "Win32_PerfFormattedData_Tcpip_TCPv4"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Tcpip_TCPv6 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataTcpipTCPv6 = "Win32_PerfFormattedData_Tcpip_TCPv6"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Tcpip_UDPv4 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataTcpipUDPv4 = "Win32_PerfFormattedData_Tcpip_UDPv4"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_Tcpip_UDPv6 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataTcpipUDPv6 = "Win32_PerfFormattedData_Tcpip_UDPv6"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_TCPIPCounters_TCPIPPerformanceDiagnostics クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataTCPIPCountersTCPIPPerformanceDiagnostics = "Win32_PerfFormattedData_TCPIPCounters_TCPIPPerformanceDiagnostics"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_TermService_TerminalServicesSession クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataTermServiceTerminalServicesSession = "Win32_PerfFormattedData_TermService_TerminalServicesSession"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_UGatherer_SearchGathererProjects クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataUGathererSearchGathererProjects = "Win32_PerfFormattedData_UGatherer_SearchGathererProjects"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_UGTHRSVC_SearchGatherer クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataUGTHRSVCSearchGatherer = "Win32_PerfFormattedData_UGTHRSVC_SearchGatherer"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_usbhub_USB クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDatausbhubUSB = "Win32_PerfFormattedData_usbhub_USB"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_VidPerfProvider_HyperVVMVidNumaNode クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataVidPerfProviderHyperVVMVidNumaNode = "Win32_PerfFormattedData_VidPerfProvider_HyperVVMVidNumaNode"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_VidPerfProvider_HyperVVMVidPartition クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataVidPerfProviderHyperVVMVidPartition = "Win32_PerfFormattedData_VidPerfProvider_HyperVVMVidPartition"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_VmbusStats_HyperVVirtualMachineBus クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataVmbusStatsHyperVVirtualMachineBus = "Win32_PerfFormattedData_VmbusStats_HyperVVirtualMachineBus"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_VmmsVirtualMachineStats_HyperVVirtualMachineHealthSummary クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataVmmsVirtualMachineStatsHyperVVirtualMachineHealthSummary = "Win32_PerfFormattedData_VmmsVirtualMachineStats_HyperVVirtualMachineHealthSummary"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_VmmsVirtualMachineStats_HyperVVirtualMachineSummary クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataVmmsVirtualMachineStatsHyperVVirtualMachineSummary = "Win32_PerfFormattedData_VmmsVirtualMachineStats_HyperVVirtualMachineSummary"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_VmTaskManagerStats_HyperVTaskManagerDetail クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataVmTaskManagerStatsHyperVTaskManagerDetail = "Win32_PerfFormattedData_VmTaskManagerStats_HyperVTaskManagerDetail"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_W3SVC_WebService クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataW3SVCWebService = "Win32_PerfFormattedData_W3SVC_WebService"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_W3SVC_WebServiceCache クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataW3SVCWebServiceCache = "Win32_PerfFormattedData_W3SVC_WebServiceCache"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_W3SVCW3WPCounterProvider_W3SVCW3WP クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataW3SVCW3WPCounterProviderW3SVCW3WP = "Win32_PerfFormattedData_W3SVCW3WPCounterProvider_W3SVCW3WP"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_WASW3WPCounterProvider_WASW3WP クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataWASW3WPCounterProviderWASW3WP = "Win32_PerfFormattedData_WASW3WPCounterProvider_WASW3WP"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_WindowsMediaPlayer_WindowsMediaPlayerMetadata クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataWindowsMediaPlayerWindowsMediaPlayerMetadata = "Win32_PerfFormattedData_WindowsMediaPlayer_WindowsMediaPlayerMetadata"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_WindowsWorkflowFoundation3000_WindowsWorkflowFoundation クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataWindowsWorkflowFoundation3000WindowsWorkflowFoundation = "Win32_PerfFormattedData_WindowsWorkflowFoundation3000_WindowsWorkflowFoundation"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_WindowsWorkflowFoundation4000_WFSystemWorkflow4000 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataWindowsWorkflowFoundation4000WFSystemWorkflow4000 = "Win32_PerfFormattedData_WindowsWorkflowFoundation4000_WFSystemWorkflow4000"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_WorkflowServiceHost4000_WorkflowServiceHost4000 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataWorkflowServiceHost4000WorkflowServiceHost4000 = "Win32_PerfFormattedData_WorkflowServiceHost4000_WorkflowServiceHost4000"

'*-----------------------------------------------------------------------------
'* Win32_PerfFormattedData_WSearchIdxPi_SearchIndexer クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfFormattedDataWSearchIdxPiSearchIndexer = "Win32_PerfFormattedData_WSearchIdxPi_SearchIndexer"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawData = "Win32_PerfRawData"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_AFDCounters_MicrosoftWinsockBSP クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataAFDCountersMicrosoftWinsockBSP = "Win32_PerfRawData_AFDCounters_MicrosoftWinsockBSP"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_APPPOOLCountersProvider_APPPOOLWAS クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataAPPPOOLCountersProviderAPPPOOLWAS = "Win32_PerfRawData_APPPOOLCountersProvider_APPPOOLWAS"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_ASP_ActiveServerPages クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataASPActiveServerPages = "Win32_PerfRawData_ASP_ActiveServerPages"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_ASPNET_ASPNET クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataASPNETASPNET = "Win32_PerfRawData_ASPNET_ASPNET"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_ASPNET_ASPNETApplications クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataASPNETASPNETApplications = "Win32_PerfRawData_ASPNET_ASPNETApplications"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_ASPNET2050727_ASPNETAppsv2050727 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataASPNET2050727ASPNETAppsv2050727 = "Win32_PerfRawData_ASPNET2050727_ASPNETAppsv2050727"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_ASPNET2050727_ASPNETv2050727 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataASPNET2050727ASPNETv2050727 = "Win32_PerfRawData_ASPNET2050727_ASPNETv2050727"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_ASPNET4030319_ASPNETAppsv4030319 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataASPNET4030319ASPNETAppsv4030319 = "Win32_PerfRawData_ASPNET4030319_ASPNETAppsv4030319"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_ASPNET4030319_ASPNETv4030319 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataASPNET4030319ASPNETv4030319 = "Win32_PerfRawData_ASPNET4030319_ASPNETv4030319"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_aspnetstate_ASPNETStateService クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataaspnetstateASPNETStateService = "Win32_PerfRawData_aspnetstate_ASPNETStateService"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_AuthorizationManager_AuthorizationManagerApplications クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataAuthorizationManagerAuthorizationManagerApplications = "Win32_PerfRawData_AuthorizationManager_AuthorizationManagerApplications"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_BalancerStats_HyperVDynamicMemoryBalancer クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataBalancerStatsHyperVDynamicMemoryBalancer = "Win32_PerfRawData_BalancerStats_HyperVDynamicMemoryBalancer"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_BalancerStats_HyperVDynamicMemoryVM クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataBalancerStatsHyperVDynamicMemoryVM = "Win32_PerfRawData_BalancerStats_HyperVDynamicMemoryVM"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_BITS_BITSNetUtilization クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataBITSBITSNetUtilization = "Win32_PerfRawData_BITS_BITSNetUtilization"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_DNS64Global クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersDNS64Global = "Win32_PerfRawData_Counters_DNS64Global"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_EventTracingforWindows クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersEventTracingforWindows = "Win32_PerfRawData_Counters_EventTracingforWindows"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_EventTracingforWindowsSession クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersEventTracingforWindowsSession = "Win32_PerfRawData_Counters_EventTracingforWindowsSession"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_FileSystemDiskActivity クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersFileSystemDiskActivity = "Win32_PerfRawData_Counters_FileSystemDiskActivity"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_GenericIKEv1AuthIPandIKEv2 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersGenericIKEv1AuthIPandIKEv2 = "Win32_PerfRawData_Counters_GenericIKEv1AuthIPandIKEv2"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_HTTPService クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersHTTPService = "Win32_PerfRawData_Counters_HTTPService"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_HTTPServiceRequestQueues クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersHTTPServiceRequestQueues = "Win32_PerfRawData_Counters_HTTPServiceRequestQueues"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_HTTPServiceUrlGroups クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersHTTPServiceUrlGroups = "Win32_PerfRawData_Counters_HTTPServiceUrlGroups"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_HyperVDynamicMemoryIntegrationService クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersHyperVDynamicMemoryIntegrationService = "Win32_PerfRawData_Counters_HyperVDynamicMemoryIntegrationService"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_HyperVVirtualMachineBusPipes クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersHyperVVirtualMachineBusPipes = "Win32_PerfRawData_Counters_HyperVVirtualMachineBusPipes"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_IPHTTPSGlobal クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersIPHTTPSGlobal = "Win32_PerfRawData_Counters_IPHTTPSGlobal"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_IPHTTPSSession クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersIPHTTPSSession = "Win32_PerfRawData_Counters_IPHTTPSSession"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_IPsecAuthIPIPv4 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersIPsecAuthIPIPv4 = "Win32_PerfRawData_Counters_IPsecAuthIPIPv4"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_IPsecAuthIPIPv6 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersIPsecAuthIPIPv6 = "Win32_PerfRawData_Counters_IPsecAuthIPIPv6"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_IPsecConnections クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersIPsecConnections = "Win32_PerfRawData_Counters_IPsecConnections"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_IPsecDoSProtection クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersIPsecDoSProtection = "Win32_PerfRawData_Counters_IPsecDoSProtection"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_IPsecDriver クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersIPsecDriver = "Win32_PerfRawData_Counters_IPsecDriver"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_IPsecIKEv1IPv4 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersIPsecIKEv1IPv4 = "Win32_PerfRawData_Counters_IPsecIKEv1IPv4"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_IPsecIKEv1IPv6 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersIPsecIKEv1IPv6 = "Win32_PerfRawData_Counters_IPsecIKEv1IPv6"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_IPsecIKEv2IPv4 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersIPsecIKEv2IPv4 = "Win32_PerfRawData_Counters_IPsecIKEv2IPv4"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_IPsecIKEv2IPv6 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersIPsecIKEv2IPv6 = "Win32_PerfRawData_Counters_IPsecIKEv2IPv6"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_Netlogon クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersNetlogon = "Win32_PerfRawData_Counters_Netlogon"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_NetworkQoSPolicy クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersNetworkQoSPolicy = "Win32_PerfRawData_Counters_NetworkQoSPolicy"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_PacerFlow クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersPacerFlow = "Win32_PerfRawData_Counters_PacerFlow"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_PacerPipe クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersPacerPipe = "Win32_PerfRawData_Counters_PacerPipe"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_PacketDirectECUtilization クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersPacketDirectECUtilization = "Win32_PerfRawData_Counters_PacketDirectECUtilization"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_PacketDirectQueueDepth クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersPacketDirectQueueDepth = "Win32_PerfRawData_Counters_PacketDirectQueueDepth"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_PacketDirectReceiveCounters クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersPacketDirectReceiveCounters = "Win32_PerfRawData_Counters_PacketDirectReceiveCounters"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_PacketDirectReceiveFilters クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersPacketDirectReceiveFilters = "Win32_PerfRawData_Counters_PacketDirectReceiveFilters"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_PacketDirectTransmitCounters クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersPacketDirectTransmitCounters = "Win32_PerfRawData_Counters_PacketDirectTransmitCounters"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_PerProcessorNetworkActivityCycles クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersPerProcessorNetworkActivityCycles = "Win32_PerfRawData_Counters_PerProcessorNetworkActivityCycles"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_PerProcessorNetworkInterfaceCardActivity クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersPerProcessorNetworkInterfaceCardActivity = "Win32_PerfRawData_Counters_PerProcessorNetworkInterfaceCardActivity"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_PhysicalNetworkInterfaceCardActivity クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersPhysicalNetworkInterfaceCardActivity = "Win32_PerfRawData_Counters_PhysicalNetworkInterfaceCardActivity"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_PowerShellWorkflow クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersPowerShellWorkflow = "Win32_PerfRawData_Counters_PowerShellWorkflow"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_ProcessorInformation クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersProcessorInformation = "Win32_PerfRawData_Counters_ProcessorInformation"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_RDMAActivity クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersRDMAActivity = "Win32_PerfRawData_Counters_RDMAActivity"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_RemoteFXGraphics クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersRemoteFXGraphics = "Win32_PerfRawData_Counters_RemoteFXGraphics"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_RemoteFXNetwork クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersRemoteFXNetwork = "Win32_PerfRawData_Counters_RemoteFXNetwork"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_SMBClientShares クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersSMBClientShares = "Win32_PerfRawData_Counters_SMBClientShares"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_SMBServer クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersSMBServer = "Win32_PerfRawData_Counters_SMBServer"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_SMBServerSessions クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersSMBServerSessions = "Win32_PerfRawData_Counters_SMBServerSessions"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_SMBServerShares クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersSMBServerShares = "Win32_PerfRawData_Counters_SMBServerShares"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_StorageSpacesTier クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersStorageSpacesTier = "Win32_PerfRawData_Counters_StorageSpacesTier"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_StorageSpacesWriteCache クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersStorageSpacesWriteCache = "Win32_PerfRawData_Counters_StorageSpacesWriteCache"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_Synchronization クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersSynchronization = "Win32_PerfRawData_Counters_Synchronization"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_SynchronizationNuma クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersSynchronizationNuma = "Win32_PerfRawData_Counters_SynchronizationNuma"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_TeredoClient クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersTeredoClient = "Win32_PerfRawData_Counters_TeredoClient"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_TeredoRelay クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersTeredoRelay = "Win32_PerfRawData_Counters_TeredoRelay"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_TeredoServer クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersTeredoServer = "Win32_PerfRawData_Counters_TeredoServer"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_ThermalZoneInformation クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersThermalZoneInformation = "Win32_PerfRawData_Counters_ThermalZoneInformation"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_WFP クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersWFP = "Win32_PerfRawData_Counters_WFP"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_WFPv4 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersWFPv4 = "Win32_PerfRawData_Counters_WFPv4"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_WFPv6 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersWFPv6 = "Win32_PerfRawData_Counters_WFPv6"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_WSManQuotaStatistics クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersWSManQuotaStatistics = "Win32_PerfRawData_Counters_WSManQuotaStatistics"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_XHCICommonBuffer クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersXHCICommonBuffer = "Win32_PerfRawData_Counters_XHCICommonBuffer"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_XHCIInterrupter クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersXHCIInterrupter = "Win32_PerfRawData_Counters_XHCIInterrupter"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Counters_XHCITransferRing クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataCountersXHCITransferRing = "Win32_PerfRawData_Counters_XHCITransferRing"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_DdmCounterProvider_RAS クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataDdmCounterProviderRAS = "Win32_PerfRawData_DdmCounterProvider_RAS"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_DeliveryOptimization_DeliveryOptimizationSwarm クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataDeliveryOptimizationDeliveryOptimizationSwarm = "Win32_PerfRawData_DeliveryOptimization_DeliveryOptimizationSwarm"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_DistributedRoutingTablePerf_DistributedRoutingTable クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataDistributedRoutingTablePerfDistributedRoutingTable = "Win32_PerfRawData_DistributedRoutingTablePerf_DistributedRoutingTable"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_ESENT_Database クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataESENTDatabase = "Win32_PerfRawData_ESENT_Database"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_ESENT_DatabaseInstances クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataESENTDatabaseInstances = "Win32_PerfRawData_ESENT_DatabaseInstances"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_ESENT_DatabaseTableClasses クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataESENTDatabaseTableClasses = "Win32_PerfRawData_ESENT_DatabaseTableClasses"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_EthernetPerfProvider_HyperVLegacyNetworkAdapter クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataEthernetPerfProviderHyperVLegacyNetworkAdapter = "Win32_PerfRawData_EthernetPerfProvider_HyperVLegacyNetworkAdapter"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_FaxService_FaxService クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataFaxServiceFaxService = "Win32_PerfRawData_FaxService_FaxService"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_ftpsvc_MicrosoftFTPService クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataftpsvcMicrosoftFTPService = "Win32_PerfRawData_ftpsvc_MicrosoftFTPService"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_GmoPerfProvider_HyperVVMSaveSnapshotandRestore クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataGmoPerfProviderHyperVVMSaveSnapshotandRestore = "Win32_PerfRawData_GmoPerfProvider_HyperVVMSaveSnapshotandRestore"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_HvStats_HyperVHypervisor クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataHvStatsHyperVHypervisor = "Win32_PerfRawData_HvStats_HyperVHypervisor"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_HvStats_HyperVHypervisorLogicalProcessor クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataHvStatsHyperVHypervisorLogicalProcessor = "Win32_PerfRawData_HvStats_HyperVHypervisorLogicalProcessor"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_HvStats_HyperVHypervisorPartition クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataHvStatsHyperVHypervisorPartition = "Win32_PerfRawData_HvStats_HyperVHypervisorPartition"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_HvStats_HyperVHypervisorRootPartition クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataHvStatsHyperVHypervisorRootPartition = "Win32_PerfRawData_HvStats_HyperVHypervisorRootPartition"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_HvStats_HyperVHypervisorRootVirtualProcessor クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataHvStatsHyperVHypervisorRootVirtualProcessor = "Win32_PerfRawData_HvStats_HyperVHypervisorRootVirtualProcessor"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_HvStats_HyperVHypervisorVirtualProcessor クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataHvStatsHyperVHypervisorVirtualProcessor = "Win32_PerfRawData_HvStats_HyperVHypervisorVirtualProcessor"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_IdePerfProvider_HyperVVirtualIDEController クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataIdePerfProviderHyperVVirtualIDEController = "Win32_PerfRawData_IdePerfProvider_HyperVVirtualIDEController"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_LocalSessionManager_TerminalServices クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataLocalSessionManagerTerminalServices = "Win32_PerfRawData_LocalSessionManager_TerminalServices"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Lsa_SecurityPerProcessStatistics クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataLsaSecurityPerProcessStatistics = "Win32_PerfRawData_Lsa_SecurityPerProcessStatistics"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Lsa_SecuritySystemWideStatistics クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataLsaSecuritySystemWideStatistics = "Win32_PerfRawData_Lsa_SecuritySystemWideStatistics"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_MicrosoftWindowsBitLockerDriverCountersProvider_BitLocker クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataMicrosoftWindowsBitLockerDriverCountersProviderBitLocker = "Win32_PerfRawData_MicrosoftWindowsBitLockerDriverCountersProvider_BitLocker"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_MicrosoftWindowsRemoteDesktopServicesRemoteFXSynth3dvsc_RemoteFXSynth3DVSCVMDevice クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataMicrosoftWindowsRemoteDesktopServicesRemoteFXSynth3dvscRemoteFXSynth3DVSCVMDevice = "Win32_PerfRawData_MicrosoftWindowsRemoteDesktopServicesRemoteFXSynth3dvsc_RemoteFXSynth3DVSCVMDevice"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_MicrosoftWindowsRemoteDesktopServicesRemoteFXSynth3dvsc_RemoteFXSynth3DVSCVMTransportChannel クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataMicrosoftWindowsRemoteDesktopServicesRemoteFXSynth3dvscRemoteFXSynth3DVSCVMTransportChannel = "Win32_PerfRawData_MicrosoftWindowsRemoteDesktopServicesRemoteFXSynth3dvsc_RemoteFXSynth3DVSCVMTransportChannel"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_MSDTC_DistributedTransactionCoordinator クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataMSDTCDistributedTransactionCoordinator = "Win32_PerfRawData_MSDTC_DistributedTransactionCoordinator"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_MSDTCBridge3000_MSDTCBridge3000 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataMSDTCBridge3000MSDTCBridge3000 = "Win32_PerfRawData_MSDTCBridge3000_MSDTCBridge3000"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_MSDTCBridge4000_MSDTCBridge4000 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataMSDTCBridge4000MSDTCBridge4000 = "Win32_PerfRawData_MSDTCBridge4000_MSDTCBridge4000"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_NETCLRData_NETCLRData クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataNETCLRDataNETCLRData = "Win32_PerfRawData_NETCLRData_NETCLRData"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_NETCLRNetworking_NETCLRNetworking クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataNETCLRNetworkingNETCLRNetworking = "Win32_PerfRawData_NETCLRNetworking_NETCLRNetworking"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_NETCLRNetworking4000_NETCLRNetworking4000 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataNETCLRNetworking4000NETCLRNetworking4000 = "Win32_PerfRawData_NETCLRNetworking4000_NETCLRNetworking4000"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_NETDataProviderforOracle_NETDataProviderforOracle クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataNETDataProviderforOracleNETDataProviderforOracle = "Win32_PerfRawData_NETDataProviderforOracle_NETDataProviderforOracle"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_NETDataProviderforSqlServer_NETDataProviderforSqlServer クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataNETDataProviderforSqlServerNETDataProviderforSqlServer = "Win32_PerfRawData_NETDataProviderforSqlServer_NETDataProviderforSqlServer"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_NETFramework_NETCLRExceptions クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataNETFrameworkNETCLRExceptions = "Win32_PerfRawData_NETFramework_NETCLRExceptions"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_NETFramework_NETCLRInterop クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataNETFrameworkNETCLRInterop = "Win32_PerfRawData_NETFramework_NETCLRInterop"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_NETFramework_NETCLRJit クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataNETFrameworkNETCLRJit = "Win32_PerfRawData_NETFramework_NETCLRJit"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_NETFramework_NETCLRLoading クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataNETFrameworkNETCLRLoading = "Win32_PerfRawData_NETFramework_NETCLRLoading"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_NETFramework_NETCLRLocksAndThreads クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataNETFrameworkNETCLRLocksAndThreads = "Win32_PerfRawData_NETFramework_NETCLRLocksAndThreads"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_NETFramework_NETCLRMemory クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataNETFrameworkNETCLRMemory = "Win32_PerfRawData_NETFramework_NETCLRMemory"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_NETFramework_NETCLRRemoting クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataNETFrameworkNETCLRRemoting = "Win32_PerfRawData_NETFramework_NETCLRRemoting"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_NETFramework_NETCLRSecurity クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataNETFrameworkNETCLRSecurity = "Win32_PerfRawData_NETFramework_NETCLRSecurity"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_NETMemoryCache40_NETMemoryCache40 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataNETMemoryCache40NETMemoryCache40 = "Win32_PerfRawData_NETMemoryCache40_NETMemoryCache40"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_NvspNicStats_HyperVVirtualNetworkAdapter クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataNvspNicStatsHyperVVirtualNetworkAdapter = "Win32_PerfRawData_NvspNicStats_HyperVVirtualNetworkAdapter"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_NvspPortStats_HyperVVirtualSwitchPort クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataNvspPortStatsHyperVVirtualSwitchPort = "Win32_PerfRawData_NvspPortStats_HyperVVirtualSwitchPort"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_NvspSwitchStats_HyperVVirtualSwitch クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataNvspSwitchStatsHyperVVirtualSwitch = "Win32_PerfRawData_NvspSwitchStats_HyperVVirtualSwitch"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_OfflineFiles_ClientSideCaching クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataOfflineFilesClientSideCaching = "Win32_PerfRawData_OfflineFiles_ClientSideCaching"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_OfflineFiles_OfflineFiles クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataOfflineFilesOfflineFiles = "Win32_PerfRawData_OfflineFiles_OfflineFiles"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_PeerDistSvc_BranchCache クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataPeerDistSvcBranchCache = "Win32_PerfRawData_PeerDistSvc_BranchCache"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_PeerNameResolutionProtocolPerf_PeerNameResolutionProtocol クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataPeerNameResolutionProtocolPerfPeerNameResolutionProtocol = "Win32_PerfRawData_PeerNameResolutionProtocolPerf_PeerNameResolutionProtocol"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_PerfDisk_LogicalDisk クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataPerfDiskLogicalDisk = "Win32_PerfRawData_PerfDisk_LogicalDisk"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_PerfDisk_PhysicalDisk クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataPerfDiskPhysicalDisk = "Win32_PerfRawData_PerfDisk_PhysicalDisk"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_PerfNet_Browser クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataPerfNetBrowser = "Win32_PerfRawData_PerfNet_Browser"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_PerfNet_Redirector クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataPerfNetRedirector = "Win32_PerfRawData_PerfNet_Redirector"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_PerfNet_Server クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataPerfNetServer = "Win32_PerfRawData_PerfNet_Server"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_PerfNet_ServerWorkQueues クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataPerfNetServerWorkQueues = "Win32_PerfRawData_PerfNet_ServerWorkQueues"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_PerfOS_Cache クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataPerfOSCache = "Win32_PerfRawData_PerfOS_Cache"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_PerfOS_Memory クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataPerfOSMemory = "Win32_PerfRawData_PerfOS_Memory"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_PerfOS_NUMANodeMemory クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataPerfOSNUMANodeMemory = "Win32_PerfRawData_PerfOS_NUMANodeMemory"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_PerfOS_Objects クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataPerfOSObjects = "Win32_PerfRawData_PerfOS_Objects"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_PerfOS_PagingFile クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataPerfOSPagingFile = "Win32_PerfRawData_PerfOS_PagingFile"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_PerfOS_Processor クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataPerfOSProcessor = "Win32_PerfRawData_PerfOS_Processor"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_PerfOS_System クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataPerfOSSystem = "Win32_PerfRawData_PerfOS_System"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_PerfProc_FullImage_Costly クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataPerfProcFullImageCostly = "Win32_PerfRawData_PerfProc_FullImage_Costly"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_PerfProc_Image_Costly クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataPerfProcImageCostly = "Win32_PerfRawData_PerfProc_Image_Costly"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_PerfProc_JobObject クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataPerfProcJobObject = "Win32_PerfRawData_PerfProc_JobObject"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_PerfProc_JobObjectDetails クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataPerfProcJobObjectDetails = "Win32_PerfRawData_PerfProc_JobObjectDetails"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_PerfProc_Process クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataPerfProcProcess = "Win32_PerfRawData_PerfProc_Process"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_PerfProc_ProcessAddressSpace_Costly クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataPerfProcProcessAddressSpaceCostly = "Win32_PerfRawData_PerfProc_ProcessAddressSpace_Costly"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_PerfProc_Thread クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataPerfProcThread = "Win32_PerfRawData_PerfProc_Thread"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_PerfProc_ThreadDetails_Costly クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataPerfProcThreadDetailsCostly = "Win32_PerfRawData_PerfProc_ThreadDetails_Costly"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_PowerMeterCounter_EnergyMeter クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataPowerMeterCounterEnergyMeter = "Win32_PerfRawData_PowerMeterCounter_EnergyMeter"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_PowerMeterCounter_PowerMeter クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataPowerMeterCounterPowerMeter = "Win32_PerfRawData_PowerMeterCounter_PowerMeter"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_rdyboost_ReadyBoostCache クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDatardyboostReadyBoostCache = "Win32_PerfRawData_rdyboost_ReadyBoostCache"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_RemoteAccess_RASPort クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataRemoteAccessRASPort = "Win32_PerfRawData_RemoteAccess_RASPort"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_RemoteAccess_RASTotal クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataRemoteAccessRASTotal = "Win32_PerfRawData_RemoteAccess_RASTotal"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_RemotePerfProvider_HyperVVMRemoting クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataRemotePerfProviderHyperVVMRemoting = "Win32_PerfRawData_RemotePerfProvider_HyperVVMRemoting"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_ServiceModel4000_ServiceModelEndpoint4000 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataServiceModel4000ServiceModelEndpoint4000 = "Win32_PerfRawData_ServiceModel4000_ServiceModelEndpoint4000"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_ServiceModel4000_ServiceModelOperation4000 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataServiceModel4000ServiceModelOperation4000 = "Win32_PerfRawData_ServiceModel4000_ServiceModelOperation4000"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_ServiceModel4000_ServiceModelService4000 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataServiceModel4000ServiceModelService4000 = "Win32_PerfRawData_ServiceModel4000_ServiceModelService4000"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_ServiceModelEndpoint3000_ServiceModelEndpoint3000 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataServiceModelEndpoint3000ServiceModelEndpoint3000 = "Win32_PerfRawData_ServiceModelEndpoint3000_ServiceModelEndpoint3000"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_ServiceModelOperation3000_ServiceModelOperation3000 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataServiceModelOperation3000ServiceModelOperation3000 = "Win32_PerfRawData_ServiceModelOperation3000_ServiceModelOperation3000"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_ServiceModelService3000_ServiceModelService3000 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataServiceModelService3000ServiceModelService3000 = "Win32_PerfRawData_ServiceModelService3000_ServiceModelService3000"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_SMSvcHost3000_SMSvcHost3000 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataSMSvcHost3000SMSvcHost3000 = "Win32_PerfRawData_SMSvcHost3000_SMSvcHost3000"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_SMSvcHost4000_SMSvcHost4000 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataSMSvcHost4000SMSvcHost4000 = "Win32_PerfRawData_SMSvcHost4000_SMSvcHost4000"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Spooler_PrintQueue クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataSpoolerPrintQueue = "Win32_PerfRawData_Spooler_PrintQueue"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_StorageStats_HyperVVirtualStorageDevice クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataStorageStatsHyperVVirtualStorageDevice = "Win32_PerfRawData_StorageStats_HyperVVirtualStorageDevice"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_TapiSrv_Telephony クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataTapiSrvTelephony = "Win32_PerfRawData_TapiSrv_Telephony"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_TBS_TBScounters クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataTBSTBScounters = "Win32_PerfRawData_TBS_TBScounters"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Tcpip_ICMP クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataTcpipICMP = "Win32_PerfRawData_Tcpip_ICMP"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Tcpip_ICMPv6 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataTcpipICMPv6 = "Win32_PerfRawData_Tcpip_ICMPv6"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Tcpip_IPv4 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataTcpipIPv4 = "Win32_PerfRawData_Tcpip_IPv4"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Tcpip_IPv6 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataTcpipIPv6 = "Win32_PerfRawData_Tcpip_IPv6"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Tcpip_NBTConnection クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataTcpipNBTConnection = "Win32_PerfRawData_Tcpip_NBTConnection"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Tcpip_NetworkAdapter クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataTcpipNetworkAdapter = "Win32_PerfRawData_Tcpip_NetworkAdapter"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Tcpip_NetworkInterface クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataTcpipNetworkInterface = "Win32_PerfRawData_Tcpip_NetworkInterface"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Tcpip_TCPv4 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataTcpipTCPv4 = "Win32_PerfRawData_Tcpip_TCPv4"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Tcpip_TCPv6 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataTcpipTCPv6 = "Win32_PerfRawData_Tcpip_TCPv6"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Tcpip_UDPv4 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataTcpipUDPv4 = "Win32_PerfRawData_Tcpip_UDPv4"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_Tcpip_UDPv6 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataTcpipUDPv6 = "Win32_PerfRawData_Tcpip_UDPv6"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_TCPIPCounters_TCPIPPerformanceDiagnostics クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataTCPIPCountersTCPIPPerformanceDiagnostics = "Win32_PerfRawData_TCPIPCounters_TCPIPPerformanceDiagnostics"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_TermService_TerminalServicesSession クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataTermServiceTerminalServicesSession = "Win32_PerfRawData_TermService_TerminalServicesSession"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_UGatherer_SearchGathererProjects クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataUGathererSearchGathererProjects = "Win32_PerfRawData_UGatherer_SearchGathererProjects"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_UGTHRSVC_SearchGatherer クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataUGTHRSVCSearchGatherer = "Win32_PerfRawData_UGTHRSVC_SearchGatherer"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_usbhub_USB クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDatausbhubUSB = "Win32_PerfRawData_usbhub_USB"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_VidPerfProvider_HyperVVMVidNumaNode クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataVidPerfProviderHyperVVMVidNumaNode = "Win32_PerfRawData_VidPerfProvider_HyperVVMVidNumaNode"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_VidPerfProvider_HyperVVMVidPartition クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataVidPerfProviderHyperVVMVidPartition = "Win32_PerfRawData_VidPerfProvider_HyperVVMVidPartition"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_VmbusStats_HyperVVirtualMachineBus クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataVmbusStatsHyperVVirtualMachineBus = "Win32_PerfRawData_VmbusStats_HyperVVirtualMachineBus"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_VmmsVirtualMachineStats_HyperVVirtualMachineHealthSummary クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataVmmsVirtualMachineStatsHyperVVirtualMachineHealthSummary = "Win32_PerfRawData_VmmsVirtualMachineStats_HyperVVirtualMachineHealthSummary"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_VmmsVirtualMachineStats_HyperVVirtualMachineSummary クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataVmmsVirtualMachineStatsHyperVVirtualMachineSummary = "Win32_PerfRawData_VmmsVirtualMachineStats_HyperVVirtualMachineSummary"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_VmTaskManagerStats_HyperVTaskManagerDetail クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataVmTaskManagerStatsHyperVTaskManagerDetail = "Win32_PerfRawData_VmTaskManagerStats_HyperVTaskManagerDetail"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_W3SVC_WebService クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataW3SVCWebService = "Win32_PerfRawData_W3SVC_WebService"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_W3SVC_WebServiceCache クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataW3SVCWebServiceCache = "Win32_PerfRawData_W3SVC_WebServiceCache"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_W3SVCW3WPCounterProvider_W3SVCW3WP クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataW3SVCW3WPCounterProviderW3SVCW3WP = "Win32_PerfRawData_W3SVCW3WPCounterProvider_W3SVCW3WP"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_WASW3WPCounterProvider_WASW3WP クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataWASW3WPCounterProviderWASW3WP = "Win32_PerfRawData_WASW3WPCounterProvider_WASW3WP"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_WindowsMediaPlayer_WindowsMediaPlayerMetadata クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataWindowsMediaPlayerWindowsMediaPlayerMetadata = "Win32_PerfRawData_WindowsMediaPlayer_WindowsMediaPlayerMetadata"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_WindowsWorkflowFoundation3000_WindowsWorkflowFoundation クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataWindowsWorkflowFoundation3000WindowsWorkflowFoundation = "Win32_PerfRawData_WindowsWorkflowFoundation3000_WindowsWorkflowFoundation"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_WindowsWorkflowFoundation4000_WFSystemWorkflow4000 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataWindowsWorkflowFoundation4000WFSystemWorkflow4000 = "Win32_PerfRawData_WindowsWorkflowFoundation4000_WFSystemWorkflow4000"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_WorkflowServiceHost4000_WorkflowServiceHost4000 クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataWorkflowServiceHost4000WorkflowServiceHost4000 = "Win32_PerfRawData_WorkflowServiceHost4000_WorkflowServiceHost4000"

'*-----------------------------------------------------------------------------
'* Win32_PerfRawData_WSearchIdxPi_SearchIndexer クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameWin32PerfRawDataWSearchIdxPiSearchIndexer = "Win32_PerfRawData_WSearchIdxPi_SearchIndexer"

'*-----------------------------------------------------------------------------
'* EventViewerConsumer クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameEventViewerConsumer = "EventViewerConsumer"

'*-----------------------------------------------------------------------------
'* NTEventlogProviderConfig クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameNTEventlogProviderConfig = "NTEventlogProviderConfig"

'*-----------------------------------------------------------------------------
'* OfficeSoftwareProtectionProduct クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameOfficeSoftwareProtectionProduct = "OfficeSoftwareProtectionProduct"

'*-----------------------------------------------------------------------------
'* OfficeSoftwareProtectionService クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameOfficeSoftwareProtectionService = "OfficeSoftwareProtectionService"

'*-----------------------------------------------------------------------------
'* OfficeSoftwareProtectionTokenActivationLicense クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameOfficeSoftwareProtectionTokenActivationLicense = "OfficeSoftwareProtectionTokenActivationLicense"

'*-----------------------------------------------------------------------------
'* RegistryEvent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameRegistryEvent = "RegistryEvent"

'*-----------------------------------------------------------------------------
'* RegistryKeyChangeEvent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameRegistryKeyChangeEvent = "RegistryKeyChangeEvent"

'*-----------------------------------------------------------------------------
'* RegistryTreeChangeEvent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameRegistryTreeChangeEvent = "RegistryTreeChangeEvent"

'*-----------------------------------------------------------------------------
'* RegistryValueChangeEvent クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameRegistryValueChangeEvent = "RegistryValueChangeEvent"

'*-----------------------------------------------------------------------------
'* ScriptingStandardConsumerSetting クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameScriptingStandardConsumerSetting = "ScriptingStandardConsumerSetting"

'*-----------------------------------------------------------------------------
'* SoftwareLicensingProduct クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameSoftwareLicensingProduct = "SoftwareLicensingProduct"

'*-----------------------------------------------------------------------------
'* SoftwareLicensingService クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameSoftwareLicensingService = "SoftwareLicensingService"

'*-----------------------------------------------------------------------------
'* SoftwareLicensingTokenActivationLicense クラス
'*
'*
'* WMI Provider :
'* UUID :
'*
'* @see
'*-----------------------------------------------------------------------------
Const WmiClassNameSoftwareLicensingTokenActivationLicense = "SoftwareLicensingTokenActivationLicense"

'******************************************************************************
'* メソッド定義
'******************************************************************************

'******************************************************************************
'* [概  要] SelectWmiClassName メソッド
'* [詳  細] Wmiの提供するClass名をWmiClassEnumを指定して取得する。
'*
'* @param selectNo WmiClassEnum
'* @return Class名
'*
'******************************************************************************
Public Function SelectWmiClassName(ByVal selectNo As WmiClassEnum) As String
    Select Case selectNo
        Case wmiEnumClassCIMAction
            SelectWmiClassName = WmiClassNameCIMAction
        Case wmiEnumClassCIMActionSequence
            SelectWmiClassName = WmiClassNameCIMActionSequence
        Case wmiEnumClassCIMActsAsSpare
            SelectWmiClassName = WmiClassNameCIMActsAsSpare
        Case wmiEnumClassCIMAdjacentSlots
            SelectWmiClassName = WmiClassNameCIMAdjacentSlots
        Case wmiEnumClassCIMAggregatePExtent
            SelectWmiClassName = WmiClassNameCIMAggregatePExtent
        Case wmiEnumClassCIMAggregatePSExtent
            SelectWmiClassName = WmiClassNameCIMAggregatePSExtent
        Case wmiEnumClassCIMAggregateRedundancyComponent
            SelectWmiClassName = WmiClassNameCIMAggregateRedundancyComponent
        Case wmiEnumClassCIMAlarmDevice
            SelectWmiClassName = WmiClassNameCIMAlarmDevice
        Case wmiEnumClassCIMAllocatedResource
            SelectWmiClassName = WmiClassNameCIMAllocatedResource
        Case wmiEnumClassCIMApplicationSystem
            SelectWmiClassName = WmiClassNameCIMApplicationSystem
        Case wmiEnumClassCIMApplicationSystemSoftwareFeature
            SelectWmiClassName = WmiClassNameCIMApplicationSystemSoftwareFeature
        Case wmiEnumClassCIMAssociatedAlarm
            SelectWmiClassName = WmiClassNameCIMAssociatedAlarm
        Case wmiEnumClassCIMAssociatedBattery
            SelectWmiClassName = WmiClassNameCIMAssociatedBattery
        Case wmiEnumClassCIMAssociatedCooling
            SelectWmiClassName = WmiClassNameCIMAssociatedCooling
        Case wmiEnumClassCIMAssociatedMemory
            SelectWmiClassName = WmiClassNameCIMAssociatedMemory
        Case wmiEnumClassCIMAssociatedProcessorMemory
            SelectWmiClassName = WmiClassNameCIMAssociatedProcessorMemory
        Case wmiEnumClassCIMAssociatedSensor
            SelectWmiClassName = WmiClassNameCIMAssociatedSensor
        Case wmiEnumClassCIMAssociatedSupplyCurrentSensor
            SelectWmiClassName = WmiClassNameCIMAssociatedSupplyCurrentSensor
        Case wmiEnumClassCIMAssociatedSupplyVoltageSensor
            SelectWmiClassName = WmiClassNameCIMAssociatedSupplyVoltageSensor
        Case wmiEnumClassCIMBasedOn
            SelectWmiClassName = WmiClassNameCIMBasedOn
        Case wmiEnumClassCIMBattery
            SelectWmiClassName = WmiClassNameCIMBattery
        Case wmiEnumClassCIMBinarySensor
            SelectWmiClassName = WmiClassNameCIMBinarySensor
        Case wmiEnumClassCIMBIOSElement
            SelectWmiClassName = WmiClassNameCIMBIOSElement
        Case wmiEnumClassCIMBIOSFeature
            SelectWmiClassName = WmiClassNameCIMBIOSFeature
        Case wmiEnumClassCIMBIOSFeatureBIOSElements
            SelectWmiClassName = WmiClassNameCIMBIOSFeatureBIOSElements
        Case wmiEnumClassCIMBIOSLoadedInNV
            SelectWmiClassName = WmiClassNameCIMBIOSLoadedInNV
        Case wmiEnumClassCIMBootOSFromFS
            SelectWmiClassName = WmiClassNameCIMBootOSFromFS
        Case wmiEnumClassCIMBootSAP
            SelectWmiClassName = WmiClassNameCIMBootSAP
        Case wmiEnumClassCIMBootService
            SelectWmiClassName = WmiClassNameCIMBootService
        Case wmiEnumClassCIMBootServiceAccessBySAP
            SelectWmiClassName = WmiClassNameCIMBootServiceAccessBySAP
        Case wmiEnumClassCIMCacheMemory
            SelectWmiClassName = WmiClassNameCIMCacheMemory
        Case wmiEnumClassCIMCard
            SelectWmiClassName = WmiClassNameCIMCard
        Case wmiEnumClassCIMCardInSlot
            SelectWmiClassName = WmiClassNameCIMCardInSlot
        Case wmiEnumClassCIMCardOnCard
            SelectWmiClassName = WmiClassNameCIMCardOnCard
        Case wmiEnumClassCIMCDROMDrive
            SelectWmiClassName = WmiClassNameCIMCDROMDrive
        Case wmiEnumClassCIMChassis
            SelectWmiClassName = WmiClassNameCIMChassis
        Case wmiEnumClassCIMChassisInRack
            SelectWmiClassName = WmiClassNameCIMChassisInRack
        Case wmiEnumClassCIMCheck
            SelectWmiClassName = WmiClassNameCIMCheck
        Case wmiEnumClassCIMChip
            SelectWmiClassName = WmiClassNameCIMChip
        Case wmiEnumClassCIMClusteringSAP
            SelectWmiClassName = WmiClassNameCIMClusteringSAP
        Case wmiEnumClassCIMClusteringService
            SelectWmiClassName = WmiClassNameCIMClusteringService
        Case wmiEnumClassCIMClusterServiceAccessBySAP
            SelectWmiClassName = WmiClassNameCIMClusterServiceAccessBySAP
        Case wmiEnumClassCIMCollectedCollections
            SelectWmiClassName = WmiClassNameCIMCollectedCollections
        Case wmiEnumClassCIMCollectedMSEs
            SelectWmiClassName = WmiClassNameCIMCollectedMSEs
        Case wmiEnumClassCIMCollectionOfMSEs
            SelectWmiClassName = WmiClassNameCIMCollectionOfMSEs
        Case wmiEnumClassCIMCollectionOfSensors
            SelectWmiClassName = WmiClassNameCIMCollectionOfSensors
        Case wmiEnumClassCIMCollectionSetting
            SelectWmiClassName = WmiClassNameCIMCollectionSetting
        Case wmiEnumClassCIMCompatibleProduct
            SelectWmiClassName = WmiClassNameCIMCompatibleProduct
        Case wmiEnumClassCIMComponent
            SelectWmiClassName = WmiClassNameCIMComponent
        Case wmiEnumClassCIMComputerSystem
            SelectWmiClassName = WmiClassNameCIMComputerSystem
        Case wmiEnumClassCIMComputerSystemDMA
            SelectWmiClassName = WmiClassNameCIMComputerSystemDMA
        Case wmiEnumClassCIMComputerSystemIRQ
            SelectWmiClassName = WmiClassNameCIMComputerSystemIRQ
        Case wmiEnumClassCIMComputerSystemMappedIO
            SelectWmiClassName = WmiClassNameCIMComputerSystemMappedIO
        Case wmiEnumClassCIMComputerSystemPackage
            SelectWmiClassName = WmiClassNameCIMComputerSystemPackage
        Case wmiEnumClassCIMComputerSystemResource
            SelectWmiClassName = WmiClassNameCIMComputerSystemResource
        Case wmiEnumClassCIMConfiguration
            SelectWmiClassName = WmiClassNameCIMConfiguration
        Case wmiEnumClassCIMConnectedTo
            SelectWmiClassName = WmiClassNameCIMConnectedTo
        Case wmiEnumClassCIMConnectorOnPackage
            SelectWmiClassName = WmiClassNameCIMConnectorOnPackage
        Case wmiEnumClassCIMContainer
            SelectWmiClassName = WmiClassNameCIMContainer
        Case wmiEnumClassCIMControlledBy
            SelectWmiClassName = WmiClassNameCIMControlledBy
        Case wmiEnumClassCIMController
            SelectWmiClassName = WmiClassNameCIMController
        Case wmiEnumClassCIMCoolingDevice
            SelectWmiClassName = WmiClassNameCIMCoolingDevice
        Case wmiEnumClassCIMCopyFileAction
            SelectWmiClassName = WmiClassNameCIMCopyFileAction
        Case wmiEnumClassCIMCreateDirectoryAction
            SelectWmiClassName = WmiClassNameCIMCreateDirectoryAction
        Case wmiEnumClassCIMCurrentSensor
            SelectWmiClassName = WmiClassNameCIMCurrentSensor
        Case wmiEnumClassCIMDataFile
            SelectWmiClassName = WmiClassNameCIMDataFile
        Case wmiEnumClassCIMDependency
            SelectWmiClassName = WmiClassNameCIMDependency
        Case wmiEnumClassCIMDependencyContext
            SelectWmiClassName = WmiClassNameCIMDependencyContext
        Case wmiEnumClassCIMDesktopMonitor
            SelectWmiClassName = WmiClassNameCIMDesktopMonitor
        Case wmiEnumClassCIMDeviceAccessedByFile
            SelectWmiClassName = WmiClassNameCIMDeviceAccessedByFile
        Case wmiEnumClassCIMDeviceConnection
            SelectWmiClassName = WmiClassNameCIMDeviceConnection
        Case wmiEnumClassCIMDeviceErrorCounts
            SelectWmiClassName = WmiClassNameCIMDeviceErrorCounts
        Case wmiEnumClassCIMDeviceFile
            SelectWmiClassName = WmiClassNameCIMDeviceFile
        Case wmiEnumClassCIMDeviceSAPImplementation
            SelectWmiClassName = WmiClassNameCIMDeviceSAPImplementation
        Case wmiEnumClassCIMDeviceServiceImplementation
            SelectWmiClassName = WmiClassNameCIMDeviceServiceImplementation
        Case wmiEnumClassCIMDeviceSoftware
            SelectWmiClassName = WmiClassNameCIMDeviceSoftware
        Case wmiEnumClassCIMDirectory
            SelectWmiClassName = WmiClassNameCIMDirectory
        Case wmiEnumClassCIMDirectoryAction
            SelectWmiClassName = WmiClassNameCIMDirectoryAction
        Case wmiEnumClassCIMDirectoryContainsFile
            SelectWmiClassName = WmiClassNameCIMDirectoryContainsFile
        Case wmiEnumClassCIMDirectorySpecification
            SelectWmiClassName = WmiClassNameCIMDirectorySpecification
        Case wmiEnumClassCIMDirectorySpecificationFile
            SelectWmiClassName = WmiClassNameCIMDirectorySpecificationFile
        Case wmiEnumClassCIMDiscreteSensor
            SelectWmiClassName = WmiClassNameCIMDiscreteSensor
        Case wmiEnumClassCIMDiskDrive
            SelectWmiClassName = WmiClassNameCIMDiskDrive
        Case wmiEnumClassCIMDisketteDrive
            SelectWmiClassName = WmiClassNameCIMDisketteDrive
        Case wmiEnumClassCIMDiskPartition
            SelectWmiClassName = WmiClassNameCIMDiskPartition
        Case wmiEnumClassCIMDiskSpaceCheck
            SelectWmiClassName = WmiClassNameCIMDiskSpaceCheck
        Case wmiEnumClassCIMDisplay
            SelectWmiClassName = WmiClassNameCIMDisplay
        Case wmiEnumClassCIMDMA
            SelectWmiClassName = WmiClassNameCIMDMA
        Case wmiEnumClassCIMDocked
            SelectWmiClassName = WmiClassNameCIMDocked
        Case wmiEnumClassCIMElementCapacity
            SelectWmiClassName = WmiClassNameCIMElementCapacity
        Case wmiEnumClassCIMElementConfiguration
            SelectWmiClassName = WmiClassNameCIMElementConfiguration
        Case wmiEnumClassCIMElementSetting
            SelectWmiClassName = WmiClassNameCIMElementSetting
        Case wmiEnumClassCIMElementsLinked
            SelectWmiClassName = WmiClassNameCIMElementsLinked
        Case wmiEnumClassCIMErrorCountersForDevice
            SelectWmiClassName = WmiClassNameCIMErrorCountersForDevice
        Case wmiEnumClassCIMExecuteProgram
            SelectWmiClassName = WmiClassNameCIMExecuteProgram
        Case wmiEnumClassCIMExport
            SelectWmiClassName = WmiClassNameCIMExport
        Case wmiEnumClassCIMExtraCapacityGroup
            SelectWmiClassName = WmiClassNameCIMExtraCapacityGroup
        Case wmiEnumClassCIMFan
            SelectWmiClassName = WmiClassNameCIMFan
        Case wmiEnumClassCIMFileAction
            SelectWmiClassName = WmiClassNameCIMFileAction
        Case wmiEnumClassCIMFileSpecification
            SelectWmiClassName = WmiClassNameCIMFileSpecification
        Case wmiEnumClassCIMFileStorage
            SelectWmiClassName = WmiClassNameCIMFileStorage
        Case wmiEnumClassCIMFileSystem
            SelectWmiClassName = WmiClassNameCIMFileSystem
        Case wmiEnumClassCIMFlatPanel
            SelectWmiClassName = WmiClassNameCIMFlatPanel
        Case wmiEnumClassCIMFromDirectoryAction
            SelectWmiClassName = WmiClassNameCIMFromDirectoryAction
        Case wmiEnumClassCIMFromDirectorySpecification
            SelectWmiClassName = WmiClassNameCIMFromDirectorySpecification
        Case wmiEnumClassCIMFRU
            SelectWmiClassName = WmiClassNameCIMFRU
        Case wmiEnumClassCIMFRUIncludesProduct
            SelectWmiClassName = WmiClassNameCIMFRUIncludesProduct
        Case wmiEnumClassCIMFRUPhysicalElements
            SelectWmiClassName = WmiClassNameCIMFRUPhysicalElements
        Case wmiEnumClassCIMHeatPipe
            SelectWmiClassName = WmiClassNameCIMHeatPipe
        Case wmiEnumClassCIMHostedAccessPoint
            SelectWmiClassName = WmiClassNameCIMHostedAccessPoint
        Case wmiEnumClassCIMHostedBootSAP
            SelectWmiClassName = WmiClassNameCIMHostedBootSAP
        Case wmiEnumClassCIMHostedBootService
            SelectWmiClassName = WmiClassNameCIMHostedBootService
        Case wmiEnumClassCIMHostedFileSystem
            SelectWmiClassName = WmiClassNameCIMHostedFileSystem
        Case wmiEnumClassCIMHostedJobDestination
            SelectWmiClassName = WmiClassNameCIMHostedJobDestination
        Case wmiEnumClassCIMHostedService
            SelectWmiClassName = WmiClassNameCIMHostedService
        Case wmiEnumClassCIMInfraredController
            SelectWmiClassName = WmiClassNameCIMInfraredController
        Case wmiEnumClassCIMInstalledOS
            SelectWmiClassName = WmiClassNameCIMInstalledOS
        Case wmiEnumClassCIMInstalledSoftwareElement
            SelectWmiClassName = WmiClassNameCIMInstalledSoftwareElement
        Case wmiEnumClassCIMIRQ
            SelectWmiClassName = WmiClassNameCIMIRQ
        Case wmiEnumClassCIMJob
            SelectWmiClassName = WmiClassNameCIMJob
        Case wmiEnumClassCIMJobDestination
            SelectWmiClassName = WmiClassNameCIMJobDestination
        Case wmiEnumClassCIMJobDestinationJobs
            SelectWmiClassName = WmiClassNameCIMJobDestinationJobs
        Case wmiEnumClassCIMKeyboard
            SelectWmiClassName = WmiClassNameCIMKeyboard
        Case wmiEnumClassCIMLinkHasConnector
            SelectWmiClassName = WmiClassNameCIMLinkHasConnector
        Case wmiEnumClassCIMLocalFileSystem
            SelectWmiClassName = WmiClassNameCIMLocalFileSystem
        Case wmiEnumClassCIMLocation
            SelectWmiClassName = WmiClassNameCIMLocation
        Case wmiEnumClassCIMLogicalDevice
            SelectWmiClassName = WmiClassNameCIMLogicalDevice
        Case wmiEnumClassCIMLogicalDisk
            SelectWmiClassName = WmiClassNameCIMLogicalDisk
        Case wmiEnumClassCIMLogicalDiskBasedOnPartition
            SelectWmiClassName = WmiClassNameCIMLogicalDiskBasedOnPartition
        Case wmiEnumClassCIMLogicalDiskBasedOnVolumeSet
            SelectWmiClassName = WmiClassNameCIMLogicalDiskBasedOnVolumeSet
        Case wmiEnumClassCIMLogicalElement
            SelectWmiClassName = WmiClassNameCIMLogicalElement
        Case wmiEnumClassCIMLogicalFile
            SelectWmiClassName = WmiClassNameCIMLogicalFile
        Case wmiEnumClassCIMLogicalIdentity
            SelectWmiClassName = WmiClassNameCIMLogicalIdentity
        Case wmiEnumClassCIMMagnetoOpticalDrive
            SelectWmiClassName = WmiClassNameCIMMagnetoOpticalDrive
        Case wmiEnumClassCIMManagedSystemElement
            SelectWmiClassName = WmiClassNameCIMManagedSystemElement
        Case wmiEnumClassCIMManagementController
            SelectWmiClassName = WmiClassNameCIMManagementController
        Case wmiEnumClassCIMMediaAccessDevice
            SelectWmiClassName = WmiClassNameCIMMediaAccessDevice
        Case wmiEnumClassCIMMediaPresent
            SelectWmiClassName = WmiClassNameCIMMediaPresent
        Case wmiEnumClassCIMMemory
            SelectWmiClassName = WmiClassNameCIMMemory
        Case wmiEnumClassCIMMemoryCapacity
            SelectWmiClassName = WmiClassNameCIMMemoryCapacity
        Case wmiEnumClassCIMMemoryCheck
            SelectWmiClassName = WmiClassNameCIMMemoryCheck
        Case wmiEnumClassCIMMemoryMappedIO
            SelectWmiClassName = WmiClassNameCIMMemoryMappedIO
        Case wmiEnumClassCIMMemoryOnCard
            SelectWmiClassName = WmiClassNameCIMMemoryOnCard
        Case wmiEnumClassCIMMemoryWithMedia
            SelectWmiClassName = WmiClassNameCIMMemoryWithMedia
        Case wmiEnumClassCIMModifySettingAction
            SelectWmiClassName = WmiClassNameCIMModifySettingAction
        Case wmiEnumClassCIMMonitorResolution
            SelectWmiClassName = WmiClassNameCIMMonitorResolution
        Case wmiEnumClassCIMMonitorSetting
            SelectWmiClassName = WmiClassNameCIMMonitorSetting
        Case wmiEnumClassCIMMount
            SelectWmiClassName = WmiClassNameCIMMount
        Case wmiEnumClassCIMMultiStateSensor
            SelectWmiClassName = WmiClassNameCIMMultiStateSensor
        Case wmiEnumClassCIMNetworkAdapter
            SelectWmiClassName = WmiClassNameCIMNetworkAdapter
        Case wmiEnumClassCIMNFS
            SelectWmiClassName = WmiClassNameCIMNFS
        Case wmiEnumClassCIMNonVolatileStorage
            SelectWmiClassName = WmiClassNameCIMNonVolatileStorage
        Case wmiEnumClassCIMNumericSensor
            SelectWmiClassName = WmiClassNameCIMNumericSensor
        Case wmiEnumClassCIMOperatingSystem
            SelectWmiClassName = WmiClassNameCIMOperatingSystem
        Case wmiEnumClassCIMOperatingSystemSoftwareFeature
            SelectWmiClassName = WmiClassNameCIMOperatingSystemSoftwareFeature
        Case wmiEnumClassCIMOSProcess
            SelectWmiClassName = WmiClassNameCIMOSProcess
        Case wmiEnumClassCIMOSVersionCheck
            SelectWmiClassName = WmiClassNameCIMOSVersionCheck
        Case wmiEnumClassCIMPackageAlarm
            SelectWmiClassName = WmiClassNameCIMPackageAlarm
        Case wmiEnumClassCIMPackageCooling
            SelectWmiClassName = WmiClassNameCIMPackageCooling
        Case wmiEnumClassCIMPackagedComponent
            SelectWmiClassName = WmiClassNameCIMPackagedComponent
        Case wmiEnumClassCIMPackageInChassis
            SelectWmiClassName = WmiClassNameCIMPackageInChassis
        Case wmiEnumClassCIMPackageInSlot
            SelectWmiClassName = WmiClassNameCIMPackageInSlot
        Case wmiEnumClassCIMPackageTempSensor
            SelectWmiClassName = WmiClassNameCIMPackageTempSensor
        Case wmiEnumClassCIMParallelController
            SelectWmiClassName = WmiClassNameCIMParallelController
        Case wmiEnumClassCIMParticipatesInSet
            SelectWmiClassName = WmiClassNameCIMParticipatesInSet
        Case wmiEnumClassCIMPCIController
            SelectWmiClassName = WmiClassNameCIMPCIController
        Case wmiEnumClassCIMPCMCIAController
            SelectWmiClassName = WmiClassNameCIMPCMCIAController
        Case wmiEnumClassCIMPCVideoController
            SelectWmiClassName = WmiClassNameCIMPCVideoController
        Case wmiEnumClassCIMPExtentRedundancyComponent
            SelectWmiClassName = WmiClassNameCIMPExtentRedundancyComponent
        Case wmiEnumClassCIMPhysicalCapacity
            SelectWmiClassName = WmiClassNameCIMPhysicalCapacity
        Case wmiEnumClassCIMPhysicalComponent
            SelectWmiClassName = WmiClassNameCIMPhysicalComponent
        Case wmiEnumClassCIMPhysicalConnector
            SelectWmiClassName = WmiClassNameCIMPhysicalConnector
        Case wmiEnumClassCIMPhysicalElement
            SelectWmiClassName = WmiClassNameCIMPhysicalElement
        Case wmiEnumClassCIMPhysicalElementLocation
            SelectWmiClassName = WmiClassNameCIMPhysicalElementLocation
        Case wmiEnumClassCIMPhysicalExtent
            SelectWmiClassName = WmiClassNameCIMPhysicalExtent
        Case wmiEnumClassCIMPhysicalFrame
            SelectWmiClassName = WmiClassNameCIMPhysicalFrame
        Case wmiEnumClassCIMPhysicalLink
            SelectWmiClassName = WmiClassNameCIMPhysicalLink
        Case wmiEnumClassCIMPhysicalMedia
            SelectWmiClassName = WmiClassNameCIMPhysicalMedia
        Case wmiEnumClassCIMPhysicalMemory
            SelectWmiClassName = WmiClassNameCIMPhysicalMemory
        Case wmiEnumClassCIMPhysicalPackage
            SelectWmiClassName = WmiClassNameCIMPhysicalPackage
        Case wmiEnumClassCIMPointingDevice
            SelectWmiClassName = WmiClassNameCIMPointingDevice
        Case wmiEnumClassCIMPotsModem
            SelectWmiClassName = WmiClassNameCIMPotsModem
        Case wmiEnumClassCIMPowerSupply
            SelectWmiClassName = WmiClassNameCIMPowerSupply
        Case wmiEnumClassCIMPrinter
            SelectWmiClassName = WmiClassNameCIMPrinter
        Case wmiEnumClassCIMProcess
            SelectWmiClassName = WmiClassNameCIMProcess
        Case wmiEnumClassCIMProcessExecutable
            SelectWmiClassName = WmiClassNameCIMProcessExecutable
        Case wmiEnumClassCIMProcessor
            SelectWmiClassName = WmiClassNameCIMProcessor
        Case wmiEnumClassCIMProcessThread
            SelectWmiClassName = WmiClassNameCIMProcessThread
        Case wmiEnumClassCIMProduct
            SelectWmiClassName = WmiClassNameCIMProduct
        Case wmiEnumClassCIMProductFRU
            SelectWmiClassName = WmiClassNameCIMProductFRU
        Case wmiEnumClassCIMProductParentChild
            SelectWmiClassName = WmiClassNameCIMProductParentChild
        Case wmiEnumClassCIMProductPhysicalElements
            SelectWmiClassName = WmiClassNameCIMProductPhysicalElements
        Case wmiEnumClassCIMProductProductDependency
            SelectWmiClassName = WmiClassNameCIMProductProductDependency
        Case wmiEnumClassCIMProductSoftwareFeatures
            SelectWmiClassName = WmiClassNameCIMProductSoftwareFeatures
        Case wmiEnumClassCIMProductSupport
            SelectWmiClassName = WmiClassNameCIMProductSupport
        Case wmiEnumClassCIMProtectedSpaceExtent
            SelectWmiClassName = WmiClassNameCIMProtectedSpaceExtent
        Case wmiEnumClassCIMPSExtentBasedOnPExtent
            SelectWmiClassName = WmiClassNameCIMPSExtentBasedOnPExtent
        Case wmiEnumClassCIMRack
            SelectWmiClassName = WmiClassNameCIMRack
        Case wmiEnumClassCIMRealizes
            SelectWmiClassName = WmiClassNameCIMRealizes
        Case wmiEnumClassCIMRealizesAggregatePExtent
            SelectWmiClassName = WmiClassNameCIMRealizesAggregatePExtent
        Case wmiEnumClassCIMRealizesDiskPartition
            SelectWmiClassName = WmiClassNameCIMRealizesDiskPartition
        Case wmiEnumClassCIMRealizesPExtent
            SelectWmiClassName = WmiClassNameCIMRealizesPExtent
        Case wmiEnumClassCIMRebootAction
            SelectWmiClassName = WmiClassNameCIMRebootAction
        Case wmiEnumClassCIMRedundancyComponent
            SelectWmiClassName = WmiClassNameCIMRedundancyComponent
        Case wmiEnumClassCIMRedundancyGroup
            SelectWmiClassName = WmiClassNameCIMRedundancyGroup
        Case wmiEnumClassCIMRefrigeration
            SelectWmiClassName = WmiClassNameCIMRefrigeration
        Case wmiEnumClassCIMRelatedStatistics
            SelectWmiClassName = WmiClassNameCIMRelatedStatistics
        Case wmiEnumClassCIMRemoteFileSystem
            SelectWmiClassName = WmiClassNameCIMRemoteFileSystem
        Case wmiEnumClassCIMRemoveDirectoryAction
            SelectWmiClassName = WmiClassNameCIMRemoveDirectoryAction
        Case wmiEnumClassCIMRemoveFileAction
            SelectWmiClassName = WmiClassNameCIMRemoveFileAction
        Case wmiEnumClassCIMReplacementSet
            SelectWmiClassName = WmiClassNameCIMReplacementSet
        Case wmiEnumClassCIMResidesOnExtent
            SelectWmiClassName = WmiClassNameCIMResidesOnExtent
        Case wmiEnumClassCIMRunningOS
            SelectWmiClassName = WmiClassNameCIMRunningOS
        Case wmiEnumClassCIMSAPSAPDependency
            SelectWmiClassName = WmiClassNameCIMSAPSAPDependency
        Case wmiEnumClassCIMScanner
            SelectWmiClassName = WmiClassNameCIMScanner
        Case wmiEnumClassCIMSCSIController
            SelectWmiClassName = WmiClassNameCIMSCSIController
        Case wmiEnumClassCIMSCSIInterface
            SelectWmiClassName = WmiClassNameCIMSCSIInterface
        Case wmiEnumClassCIMSensor
            SelectWmiClassName = WmiClassNameCIMSensor
        Case wmiEnumClassCIMSerialController
            SelectWmiClassName = WmiClassNameCIMSerialController
        Case wmiEnumClassCIMSerialInterface
            SelectWmiClassName = WmiClassNameCIMSerialInterface
        Case wmiEnumClassCIMService
            SelectWmiClassName = WmiClassNameCIMService
        Case wmiEnumClassCIMServiceAccessBySAP
            SelectWmiClassName = WmiClassNameCIMServiceAccessBySAP
        Case wmiEnumClassCIMServiceAccessPoint
            SelectWmiClassName = WmiClassNameCIMServiceAccessPoint
        Case wmiEnumClassCIMServiceSAPDependency
            SelectWmiClassName = WmiClassNameCIMServiceSAPDependency
        Case wmiEnumClassCIMServiceServiceDependency
            SelectWmiClassName = WmiClassNameCIMServiceServiceDependency
        Case wmiEnumClassCIMSetting
            SelectWmiClassName = WmiClassNameCIMSetting
        Case wmiEnumClassCIMSettingCheck
            SelectWmiClassName = WmiClassNameCIMSettingCheck
        Case wmiEnumClassCIMSettingContext
            SelectWmiClassName = WmiClassNameCIMSettingContext
        Case wmiEnumClassCIMSlot
            SelectWmiClassName = WmiClassNameCIMSlot
        Case wmiEnumClassCIMSlotInSlot
            SelectWmiClassName = WmiClassNameCIMSlotInSlot
        Case wmiEnumClassCIMSoftwareElement
            SelectWmiClassName = WmiClassNameCIMSoftwareElement
        Case wmiEnumClassCIMSoftwareElementActions
            SelectWmiClassName = WmiClassNameCIMSoftwareElementActions
        Case wmiEnumClassCIMSoftwareElementChecks
            SelectWmiClassName = WmiClassNameCIMSoftwareElementChecks
        Case wmiEnumClassCIMSoftwareElementVersionCheck
            SelectWmiClassName = WmiClassNameCIMSoftwareElementVersionCheck
        Case wmiEnumClassCIMSoftwareFeature
            SelectWmiClassName = WmiClassNameCIMSoftwareFeature
        Case wmiEnumClassCIMSoftwareFeatureSAPImplementation
            SelectWmiClassName = WmiClassNameCIMSoftwareFeatureSAPImplementation
        Case wmiEnumClassCIMSoftwareFeatureServiceImplementation
            SelectWmiClassName = WmiClassNameCIMSoftwareFeatureServiceImplementation
        Case wmiEnumClassCIMSoftwareFeatureSoftwareElements
            SelectWmiClassName = WmiClassNameCIMSoftwareFeatureSoftwareElements
        Case wmiEnumClassCIMSpareGroup
            SelectWmiClassName = WmiClassNameCIMSpareGroup
        Case wmiEnumClassCIMStatisticalInformation
            SelectWmiClassName = WmiClassNameCIMStatisticalInformation
        Case wmiEnumClassCIMStatistics
            SelectWmiClassName = WmiClassNameCIMStatistics
        Case wmiEnumClassCIMStorageDefect
            SelectWmiClassName = WmiClassNameCIMStorageDefect
        Case wmiEnumClassCIMStorageError
            SelectWmiClassName = WmiClassNameCIMStorageError
        Case wmiEnumClassCIMStorageExtent
            SelectWmiClassName = WmiClassNameCIMStorageExtent
        Case wmiEnumClassCIMStorageRedundancyGroup
            SelectWmiClassName = WmiClassNameCIMStorageRedundancyGroup
        Case wmiEnumClassCIMSupportAccess
            SelectWmiClassName = WmiClassNameCIMSupportAccess
        Case wmiEnumClassCIMSwapSpaceCheck
            SelectWmiClassName = WmiClassNameCIMSwapSpaceCheck
        Case wmiEnumClassCIMSystem
            SelectWmiClassName = WmiClassNameCIMSystem
        Case wmiEnumClassCIMSystemComponent
            SelectWmiClassName = WmiClassNameCIMSystemComponent
        Case wmiEnumClassCIMSystemDevice
            SelectWmiClassName = WmiClassNameCIMSystemDevice
        Case wmiEnumClassCIMSystemResource
            SelectWmiClassName = WmiClassNameCIMSystemResource
        Case wmiEnumClassCIMTachometer
            SelectWmiClassName = WmiClassNameCIMTachometer
        Case wmiEnumClassCIMTapeDrive
            SelectWmiClassName = WmiClassNameCIMTapeDrive
        Case wmiEnumClassCIMTemperatureSensor
            SelectWmiClassName = WmiClassNameCIMTemperatureSensor
        Case wmiEnumClassCIMThread
            SelectWmiClassName = WmiClassNameCIMThread
        Case wmiEnumClassCIMToDirectoryAction
            SelectWmiClassName = WmiClassNameCIMToDirectoryAction
        Case wmiEnumClassCIMToDirectorySpecification
            SelectWmiClassName = WmiClassNameCIMToDirectorySpecification
        Case wmiEnumClassCIMUninterruptiblePowerSupply
            SelectWmiClassName = WmiClassNameCIMUninterruptiblePowerSupply
        Case wmiEnumClassCIMUnitaryComputerSystem
            SelectWmiClassName = WmiClassNameCIMUnitaryComputerSystem
        Case wmiEnumClassCIMUSBController
            SelectWmiClassName = WmiClassNameCIMUSBController
        Case wmiEnumClassCIMUSBControllerHasHub
            SelectWmiClassName = WmiClassNameCIMUSBControllerHasHub
        Case wmiEnumClassCIMUserDevice
            SelectWmiClassName = WmiClassNameCIMUserDevice
        Case wmiEnumClassCIMVersionCompatibilityCheck
            SelectWmiClassName = WmiClassNameCIMVersionCompatibilityCheck
        Case wmiEnumClassCIMVideoBIOSElement
            SelectWmiClassName = WmiClassNameCIMVideoBIOSElement
        Case wmiEnumClassCIMVideoBIOSFeature
            SelectWmiClassName = WmiClassNameCIMVideoBIOSFeature
        Case wmiEnumClassCIMVideoBIOSFeatureVideoBIOSElements
            SelectWmiClassName = WmiClassNameCIMVideoBIOSFeatureVideoBIOSElements
        Case wmiEnumClassCIMVideoController
            SelectWmiClassName = WmiClassNameCIMVideoController
        Case wmiEnumClassCIMVideoControllerResolution
            SelectWmiClassName = WmiClassNameCIMVideoControllerResolution
        Case wmiEnumClassCIMVideoSetting
            SelectWmiClassName = WmiClassNameCIMVideoSetting
        Case wmiEnumClassCIMVolatileStorage
            SelectWmiClassName = WmiClassNameCIMVolatileStorage
        Case wmiEnumClassCIMVoltageSensor
            SelectWmiClassName = WmiClassNameCIMVoltageSensor
        Case wmiEnumClassCIMVolumeSet
            SelectWmiClassName = WmiClassNameCIMVolumeSet
        Case wmiEnumClassCIMWORMDrive
            SelectWmiClassName = WmiClassNameCIMWORMDrive
        Case wmiEnumClassMSFTNCProvAccessCheck
            SelectWmiClassName = WmiClassNameMSFTNCProvAccessCheck
        Case wmiEnumClassMSFTNCProvCancelQuery
            SelectWmiClassName = WmiClassNameMSFTNCProvCancelQuery
        Case wmiEnumClassMSFTNCProvClientConnected
            SelectWmiClassName = WmiClassNameMSFTNCProvClientConnected
        Case wmiEnumClassMSFTNCProvEvent
            SelectWmiClassName = WmiClassNameMSFTNCProvEvent
        Case wmiEnumClassMSFTNCProvNewQuery
            SelectWmiClassName = WmiClassNameMSFTNCProvNewQuery
        Case wmiEnumClassMSFTNetBadAccount
            SelectWmiClassName = WmiClassNameMSFTNetBadAccount
        Case wmiEnumClassMSFTNetBadServiceState
            SelectWmiClassName = WmiClassNameMSFTNetBadServiceState
        Case wmiEnumClassMSFTNetBootSystemDriversFailed
            SelectWmiClassName = WmiClassNameMSFTNetBootSystemDriversFailed
        Case wmiEnumClassMSFTNetCallToFunctionFailed
            SelectWmiClassName = WmiClassNameMSFTNetCallToFunctionFailed
        Case wmiEnumClassMSFTNetCallToFunctionFailedII
            SelectWmiClassName = WmiClassNameMSFTNetCallToFunctionFailedII
        Case wmiEnumClassMSFTNetCircularDependencyAuto
            SelectWmiClassName = WmiClassNameMSFTNetCircularDependencyAuto
        Case wmiEnumClassMSFTNetCircularDependencyDemand
            SelectWmiClassName = WmiClassNameMSFTNetCircularDependencyDemand
        Case wmiEnumClassMSFTNetConnectionTimeout
            SelectWmiClassName = WmiClassNameMSFTNetConnectionTimeout
        Case wmiEnumClassMSFTNetDependOnLaterGroup
            SelectWmiClassName = WmiClassNameMSFTNetDependOnLaterGroup
        Case wmiEnumClassMSFTNetDependOnLaterService
            SelectWmiClassName = WmiClassNameMSFTNetDependOnLaterService
        Case wmiEnumClassMSFTNetFirstLogonFailed
            SelectWmiClassName = WmiClassNameMSFTNetFirstLogonFailed
        Case wmiEnumClassMSFTNetFirstLogonFailedII
            SelectWmiClassName = WmiClassNameMSFTNetFirstLogonFailedII
        Case wmiEnumClassMSFTNetReadfileTimeout
            SelectWmiClassName = WmiClassNameMSFTNetReadfileTimeout
        Case wmiEnumClassMSFTNetRevertedToLastKnownGood
            SelectWmiClassName = WmiClassNameMSFTNetRevertedToLastKnownGood
        Case wmiEnumClassMSFTNetServiceConfigBackoutFailed
            SelectWmiClassName = WmiClassNameMSFTNetServiceConfigBackoutFailed
        Case wmiEnumClassMSFTNetServiceControlSuccess
            SelectWmiClassName = WmiClassNameMSFTNetServiceControlSuccess
        Case wmiEnumClassMSFTNetServiceCrash
            SelectWmiClassName = WmiClassNameMSFTNetServiceCrash
        Case wmiEnumClassMSFTNetServiceCrashNoAction
            SelectWmiClassName = WmiClassNameMSFTNetServiceCrashNoAction
        Case wmiEnumClassMSFTNetServiceDifferentPIDConnected
            SelectWmiClassName = WmiClassNameMSFTNetServiceDifferentPIDConnected
        Case wmiEnumClassMSFTNetServiceExitFailed
            SelectWmiClassName = WmiClassNameMSFTNetServiceExitFailed
        Case wmiEnumClassMSFTNetServiceExitFailedSpecific
            SelectWmiClassName = WmiClassNameMSFTNetServiceExitFailedSpecific
        Case wmiEnumClassMSFTNetServiceLogonTypeNotGranted
            SelectWmiClassName = WmiClassNameMSFTNetServiceLogonTypeNotGranted
        Case wmiEnumClassMSFTNetServiceNotInteractive
            SelectWmiClassName = WmiClassNameMSFTNetServiceNotInteractive
        Case wmiEnumClassMSFTNetServiceRecoveryFailed
            SelectWmiClassName = WmiClassNameMSFTNetServiceRecoveryFailed
        Case wmiEnumClassMSFTNetServiceShutdownFailed
            SelectWmiClassName = WmiClassNameMSFTNetServiceShutdownFailed
        Case wmiEnumClassMSFTNetServiceSlowStartup
            SelectWmiClassName = WmiClassNameMSFTNetServiceSlowStartup
        Case wmiEnumClassMSFTNetServiceStartFailed
            SelectWmiClassName = WmiClassNameMSFTNetServiceStartFailed
        Case wmiEnumClassMSFTNetServiceStartFailedGroup
            SelectWmiClassName = WmiClassNameMSFTNetServiceStartFailedGroup
        Case wmiEnumClassMSFTNetServiceStartFailedII
            SelectWmiClassName = WmiClassNameMSFTNetServiceStartFailedII
        Case wmiEnumClassMSFTNetServiceStartFailedNone
            SelectWmiClassName = WmiClassNameMSFTNetServiceStartFailedNone
        Case wmiEnumClassMSFTNetServiceStartHung
            SelectWmiClassName = WmiClassNameMSFTNetServiceStartHung
        Case wmiEnumClassMSFTNetServiceStartTypeChanged
            SelectWmiClassName = WmiClassNameMSFTNetServiceStartTypeChanged
        Case wmiEnumClassMSFTNetServiceStatusSuccess
            SelectWmiClassName = WmiClassNameMSFTNetServiceStatusSuccess
        Case wmiEnumClassMSFTNetServiceStopControlSuccess
            SelectWmiClassName = WmiClassNameMSFTNetServiceStopControlSuccess
        Case wmiEnumClassMSFTNetSevereServiceFailed
            SelectWmiClassName = WmiClassNameMSFTNetSevereServiceFailed
        Case wmiEnumClassMSFTNetTakeOwnership
            SelectWmiClassName = WmiClassNameMSFTNetTakeOwnership
        Case wmiEnumClassMSFTNetTransactInvalid
            SelectWmiClassName = WmiClassNameMSFTNetTransactInvalid
        Case wmiEnumClassMSFTNetTransactTimeout
            SelectWmiClassName = WmiClassNameMSFTNetTransactTimeout
        Case wmiEnumClassMsftProviders
            SelectWmiClassName = WmiClassNameMsftProviders
        Case wmiEnumClassMSFTSCMEvent
            SelectWmiClassName = WmiClassNameMSFTSCMEvent
        Case wmiEnumClassMSFTSCMEventLogEvent
            SelectWmiClassName = WmiClassNameMSFTSCMEventLogEvent
        Case wmiEnumClassMSFTWMIGenericNonCOMEvent
            SelectWmiClassName = WmiClassNameMSFTWMIGenericNonCOMEvent
        Case wmiEnumClassMSFTWmiCancelNotificationSink
            SelectWmiClassName = WmiClassNameMSFTWmiCancelNotificationSink
        Case wmiEnumClassMSFTWmiConsumerProviderEvent
            SelectWmiClassName = WmiClassNameMSFTWmiConsumerProviderEvent
        Case wmiEnumClassMSFTWmiConsumerProviderLoaded
            SelectWmiClassName = WmiClassNameMSFTWmiConsumerProviderLoaded
        Case wmiEnumClassMSFTWmiConsumerProviderSinkLoaded
            SelectWmiClassName = WmiClassNameMSFTWmiConsumerProviderSinkLoaded
        Case wmiEnumClassMSFTWmiConsumerProviderSinkUnloaded
            SelectWmiClassName = WmiClassNameMSFTWmiConsumerProviderSinkUnloaded
        Case wmiEnumClassMSFTWmiConsumerProviderUnloaded
            SelectWmiClassName = WmiClassNameMSFTWmiConsumerProviderUnloaded
        Case wmiEnumClassMSFTWmiEssEvent
            SelectWmiClassName = WmiClassNameMSFTWmiEssEvent
        Case wmiEnumClassMSFTWmiFilterActivated
            SelectWmiClassName = WmiClassNameMSFTWmiFilterActivated
        Case wmiEnumClassMSFTWmiFilterDeactivated
            SelectWmiClassName = WmiClassNameMSFTWmiFilterDeactivated
        Case wmiEnumClassMSFTWmiFilterEvent
            SelectWmiClassName = WmiClassNameMSFTWmiFilterEvent
        Case wmiEnumClassMsftWmiProviderAccessCheckPost
            SelectWmiClassName = WmiClassNameMsftWmiProviderAccessCheckPost
        Case wmiEnumClassMsftWmiProviderAccessCheckPre
            SelectWmiClassName = WmiClassNameMsftWmiProviderAccessCheckPre
        Case wmiEnumClassMsftWmiProviderCancelQueryPost
            SelectWmiClassName = WmiClassNameMsftWmiProviderCancelQueryPost
        Case wmiEnumClassMsftWmiProviderCancelQueryPre
            SelectWmiClassName = WmiClassNameMsftWmiProviderCancelQueryPre
        Case wmiEnumClassMsftWmiProviderComServerLoadOperationEvent
            SelectWmiClassName = WmiClassNameMsftWmiProviderComServerLoadOperationEvent
        Case wmiEnumClassMsftWmiProviderComServerLoadOperationFailureEvent
            SelectWmiClassName = WmiClassNameMsftWmiProviderComServerLoadOperationFailureEvent
        Case wmiEnumClassMsftWmiProviderCounters
            SelectWmiClassName = WmiClassNameMsftWmiProviderCounters
        Case wmiEnumClassMsftWmiProviderCreateClassEnumAsyncEventPost
            SelectWmiClassName = WmiClassNameMsftWmiProviderCreateClassEnumAsyncEventPost
        Case wmiEnumClassMsftWmiProviderCreateClassEnumAsyncEventPre
            SelectWmiClassName = WmiClassNameMsftWmiProviderCreateClassEnumAsyncEventPre
        Case wmiEnumClassMsftWmiProviderCreateInstanceEnumAsyncEventPost
            SelectWmiClassName = WmiClassNameMsftWmiProviderCreateInstanceEnumAsyncEventPost
        Case wmiEnumClassMsftWmiProviderCreateInstanceEnumAsyncEventPre
            SelectWmiClassName = WmiClassNameMsftWmiProviderCreateInstanceEnumAsyncEventPre
        Case wmiEnumClassMsftWmiProviderDeleteClassAsyncEventPost
            SelectWmiClassName = WmiClassNameMsftWmiProviderDeleteClassAsyncEventPost
        Case wmiEnumClassMsftWmiProviderDeleteClassAsyncEventPre
            SelectWmiClassName = WmiClassNameMsftWmiProviderDeleteClassAsyncEventPre
        Case wmiEnumClassMsftWmiProviderDeleteInstanceAsyncEventPost
            SelectWmiClassName = WmiClassNameMsftWmiProviderDeleteInstanceAsyncEventPost
        Case wmiEnumClassMsftWmiProviderDeleteInstanceAsyncEventPre
            SelectWmiClassName = WmiClassNameMsftWmiProviderDeleteInstanceAsyncEventPre
        Case wmiEnumClassMsftWmiProviderExecMethodAsyncEventPost
            SelectWmiClassName = WmiClassNameMsftWmiProviderExecMethodAsyncEventPost
        Case wmiEnumClassMsftWmiProviderExecMethodAsyncEventPre
            SelectWmiClassName = WmiClassNameMsftWmiProviderExecMethodAsyncEventPre
        Case wmiEnumClassMsftWmiProviderExecQueryAsyncEventPost
            SelectWmiClassName = WmiClassNameMsftWmiProviderExecQueryAsyncEventPost
        Case wmiEnumClassMsftWmiProviderExecQueryAsyncEventPre
            SelectWmiClassName = WmiClassNameMsftWmiProviderExecQueryAsyncEventPre
        Case wmiEnumClassMsftWmiProviderGetObjectAsyncEventPost
            SelectWmiClassName = WmiClassNameMsftWmiProviderGetObjectAsyncEventPost
        Case wmiEnumClassMsftWmiProviderGetObjectAsyncEventPre
            SelectWmiClassName = WmiClassNameMsftWmiProviderGetObjectAsyncEventPre
        Case wmiEnumClassMsftWmiProviderInitializationOperationEvent
            SelectWmiClassName = WmiClassNameMsftWmiProviderInitializationOperationEvent
        Case wmiEnumClassMsftWmiProviderInitializationOperationFailureEvent
            SelectWmiClassName = WmiClassNameMsftWmiProviderInitializationOperationFailureEvent
        Case wmiEnumClassMsftWmiProviderLoadOperationEvent
            SelectWmiClassName = WmiClassNameMsftWmiProviderLoadOperationEvent
        Case wmiEnumClassMsftWmiProviderLoadOperationFailureEvent
            SelectWmiClassName = WmiClassNameMsftWmiProviderLoadOperationFailureEvent
        Case wmiEnumClassMsftWmiProviderNewQueryPost
            SelectWmiClassName = WmiClassNameMsftWmiProviderNewQueryPost
        Case wmiEnumClassMsftWmiProviderNewQueryPre
            SelectWmiClassName = WmiClassNameMsftWmiProviderNewQueryPre
        Case wmiEnumClassMsftWmiProviderOperationEvent
            SelectWmiClassName = WmiClassNameMsftWmiProviderOperationEvent
        Case wmiEnumClassMsftWmiProviderOperationEventPost
            SelectWmiClassName = WmiClassNameMsftWmiProviderOperationEventPost
        Case wmiEnumClassMsftWmiProviderOperationEventPre
            SelectWmiClassName = WmiClassNameMsftWmiProviderOperationEventPre
        Case wmiEnumClassMsftWmiProviderProvideEventsPost
            SelectWmiClassName = WmiClassNameMsftWmiProviderProvideEventsPost
        Case wmiEnumClassMsftWmiProviderProvideEventsPre
            SelectWmiClassName = WmiClassNameMsftWmiProviderProvideEventsPre
        Case wmiEnumClassMsftWmiProviderPutClassAsyncEventPost
            SelectWmiClassName = WmiClassNameMsftWmiProviderPutClassAsyncEventPost
        Case wmiEnumClassMsftWmiProviderPutClassAsyncEventPre
            SelectWmiClassName = WmiClassNameMsftWmiProviderPutClassAsyncEventPre
        Case wmiEnumClassMsftWmiProviderPutInstanceAsyncEventPost
            SelectWmiClassName = WmiClassNameMsftWmiProviderPutInstanceAsyncEventPost
        Case wmiEnumClassMsftWmiProviderPutInstanceAsyncEventPre
            SelectWmiClassName = WmiClassNameMsftWmiProviderPutInstanceAsyncEventPre
        Case wmiEnumClassMsftWmiProviderUnLoadOperationEvent
            SelectWmiClassName = WmiClassNameMsftWmiProviderUnLoadOperationEvent
        Case wmiEnumClassMSFTWmiProviderEvent
            SelectWmiClassName = WmiClassNameMSFTWmiProviderEvent
        Case wmiEnumClassMSFTWmiRegisterNotificationSink
            SelectWmiClassName = WmiClassNameMSFTWmiRegisterNotificationSink
        Case wmiEnumClassMSFTWmiSelfEvent
            SelectWmiClassName = WmiClassNameMSFTWmiSelfEvent
        Case wmiEnumClassMSFTWmiThreadPoolEvent
            SelectWmiClassName = WmiClassNameMSFTWmiThreadPoolEvent
        Case wmiEnumClassMSFTWmiThreadPoolThreadCreated
            SelectWmiClassName = WmiClassNameMSFTWmiThreadPoolThreadCreated
        Case wmiEnumClassMSFTWmiThreadPoolThreadDeleted
            SelectWmiClassName = WmiClassNameMSFTWmiThreadPoolThreadDeleted
        Case wmiEnumClassWin321394Controller
            SelectWmiClassName = WmiClassNameWin321394Controller
        Case wmiEnumClassWin321394ControllerDevice
            SelectWmiClassName = WmiClassNameWin321394ControllerDevice
        Case wmiEnumClassWin32Account
            SelectWmiClassName = WmiClassNameWin32Account
        Case wmiEnumClassWin32AccountSID
            SelectWmiClassName = WmiClassNameWin32AccountSID
        Case wmiEnumClassWin32ACE
            SelectWmiClassName = WmiClassNameWin32ACE
        Case wmiEnumClassWin32ActionCheck
            SelectWmiClassName = WmiClassNameWin32ActionCheck
        Case wmiEnumClassWin32ActiveRoute
            SelectWmiClassName = WmiClassNameWin32ActiveRoute
        Case wmiEnumClassWin32AllocatedResource
            SelectWmiClassName = WmiClassNameWin32AllocatedResource
        Case wmiEnumClassWin32ApplicationCommandLine
            SelectWmiClassName = WmiClassNameWin32ApplicationCommandLine
        Case wmiEnumClassWin32ApplicationService
            SelectWmiClassName = WmiClassNameWin32ApplicationService
        Case wmiEnumClassWin32AssociatedProcessorMemory
            SelectWmiClassName = WmiClassNameWin32AssociatedProcessorMemory
        Case wmiEnumClassWin32AutochkSetting
            SelectWmiClassName = WmiClassNameWin32AutochkSetting
        Case wmiEnumClassWin32BaseBoard
            SelectWmiClassName = WmiClassNameWin32BaseBoard
        Case wmiEnumClassWin32BaseService
            SelectWmiClassName = WmiClassNameWin32BaseService
        Case wmiEnumClassWin32Battery
            SelectWmiClassName = WmiClassNameWin32Battery
        Case wmiEnumClassWin32Binary
            SelectWmiClassName = WmiClassNameWin32Binary
        Case wmiEnumClassWin32BindImageAction
            SelectWmiClassName = WmiClassNameWin32BindImageAction
        Case wmiEnumClassWin32BIOS
            SelectWmiClassName = WmiClassNameWin32BIOS
        Case wmiEnumClassWin32BootConfiguration
            SelectWmiClassName = WmiClassNameWin32BootConfiguration
        Case wmiEnumClassWin32Bus
            SelectWmiClassName = WmiClassNameWin32Bus
        Case wmiEnumClassWin32CacheMemory
            SelectWmiClassName = WmiClassNameWin32CacheMemory
        Case wmiEnumClassWin32CDROMDrive
            SelectWmiClassName = WmiClassNameWin32CDROMDrive
        Case wmiEnumClassWin32CheckCheck
            SelectWmiClassName = WmiClassNameWin32CheckCheck
        Case wmiEnumClassWin32CIMLogicalDeviceCIMDataFile
            SelectWmiClassName = WmiClassNameWin32CIMLogicalDeviceCIMDataFile
        Case wmiEnumClassWin32ClassicCOMApplicationClasses
            SelectWmiClassName = WmiClassNameWin32ClassicCOMApplicationClasses
        Case wmiEnumClassWin32ClassicCOMClass
            SelectWmiClassName = WmiClassNameWin32ClassicCOMClass
        Case wmiEnumClassWin32ClassicCOMClassSetting
            SelectWmiClassName = WmiClassNameWin32ClassicCOMClassSetting
        Case wmiEnumClassWin32ClassicCOMClassSettings
            SelectWmiClassName = WmiClassNameWin32ClassicCOMClassSettings
        Case wmiEnumClassWin32ClassInfoAction
            SelectWmiClassName = WmiClassNameWin32ClassInfoAction
        Case wmiEnumClassWin32ClientApplicationSetting
            SelectWmiClassName = WmiClassNameWin32ClientApplicationSetting
        Case wmiEnumClassWin32ClusterShare
            SelectWmiClassName = WmiClassNameWin32ClusterShare
        Case wmiEnumClassWin32CodecFile
            SelectWmiClassName = WmiClassNameWin32CodecFile
        Case wmiEnumClassWin32CollectionStatistics
            SelectWmiClassName = WmiClassNameWin32CollectionStatistics
        Case wmiEnumClassWin32COMApplication
            SelectWmiClassName = WmiClassNameWin32COMApplication
        Case wmiEnumClassWin32COMApplicationClasses
            SelectWmiClassName = WmiClassNameWin32COMApplicationClasses
        Case wmiEnumClassWin32COMApplicationSettings
            SelectWmiClassName = WmiClassNameWin32COMApplicationSettings
        Case wmiEnumClassWin32COMClass
            SelectWmiClassName = WmiClassNameWin32COMClass
        Case wmiEnumClassWin32ComClassAutoEmulator
            SelectWmiClassName = WmiClassNameWin32ComClassAutoEmulator
        Case wmiEnumClassWin32ComClassEmulator
            SelectWmiClassName = WmiClassNameWin32ComClassEmulator
        Case wmiEnumClassWin32CommandLineAccess
            SelectWmiClassName = WmiClassNameWin32CommandLineAccess
        Case wmiEnumClassWin32ComponentCategory
            SelectWmiClassName = WmiClassNameWin32ComponentCategory
        Case wmiEnumClassWin32ComputerShutdownEvent
            SelectWmiClassName = WmiClassNameWin32ComputerShutdownEvent
        Case wmiEnumClassWin32ComputerSystem
            SelectWmiClassName = WmiClassNameWin32ComputerSystem
        Case wmiEnumClassWin32ComputerSystemEvent
            SelectWmiClassName = WmiClassNameWin32ComputerSystemEvent
        Case wmiEnumClassWin32ComputerSystemProcessor
            SelectWmiClassName = WmiClassNameWin32ComputerSystemProcessor
        Case wmiEnumClassWin32ComputerSystemProduct
            SelectWmiClassName = WmiClassNameWin32ComputerSystemProduct
        Case wmiEnumClassWin32COMSetting
            SelectWmiClassName = WmiClassNameWin32COMSetting
        Case wmiEnumClassWin32Condition
            SelectWmiClassName = WmiClassNameWin32Condition
        Case wmiEnumClassWin32ConnectionShare
            SelectWmiClassName = WmiClassNameWin32ConnectionShare
        Case wmiEnumClassWin32ControllerHasHub
            SelectWmiClassName = WmiClassNameWin32ControllerHasHub
        Case wmiEnumClassWin32CreateFolderAction
            SelectWmiClassName = WmiClassNameWin32CreateFolderAction
        Case wmiEnumClassWin32CurrentProbe
            SelectWmiClassName = WmiClassNameWin32CurrentProbe
        Case wmiEnumClassWin32CurrentTime
            SelectWmiClassName = WmiClassNameWin32CurrentTime
        Case wmiEnumClassWin32DCOMApplication
            SelectWmiClassName = WmiClassNameWin32DCOMApplication
        Case wmiEnumClassWin32DCOMApplicationAccessAllowedSetting
            SelectWmiClassName = WmiClassNameWin32DCOMApplicationAccessAllowedSetting
        Case wmiEnumClassWin32DCOMApplicationLaunchAllowedSetting
            SelectWmiClassName = WmiClassNameWin32DCOMApplicationLaunchAllowedSetting
        Case wmiEnumClassWin32DCOMApplicationSetting
            SelectWmiClassName = WmiClassNameWin32DCOMApplicationSetting
        Case wmiEnumClassWin32DependentService
            SelectWmiClassName = WmiClassNameWin32DependentService
        Case wmiEnumClassWin32Desktop
            SelectWmiClassName = WmiClassNameWin32Desktop
        Case wmiEnumClassWin32DesktopMonitor
            SelectWmiClassName = WmiClassNameWin32DesktopMonitor
        Case wmiEnumClassWin32DeviceBus
            SelectWmiClassName = WmiClassNameWin32DeviceBus
        Case wmiEnumClassWin32DeviceChangeEvent
            SelectWmiClassName = WmiClassNameWin32DeviceChangeEvent
        Case wmiEnumClassWin32DeviceMemoryAddress
            SelectWmiClassName = WmiClassNameWin32DeviceMemoryAddress
        Case wmiEnumClassWin32DeviceSettings
            SelectWmiClassName = WmiClassNameWin32DeviceSettings
        Case wmiEnumClassWin32DfsNode
            SelectWmiClassName = WmiClassNameWin32DfsNode
        Case wmiEnumClassWin32DfsNodeTarget
            SelectWmiClassName = WmiClassNameWin32DfsNodeTarget
        Case wmiEnumClassWin32DfsTarget
            SelectWmiClassName = WmiClassNameWin32DfsTarget
        Case wmiEnumClassWin32Directory
            SelectWmiClassName = WmiClassNameWin32Directory
        Case wmiEnumClassWin32DirectorySpecification
            SelectWmiClassName = WmiClassNameWin32DirectorySpecification
        Case wmiEnumClassWin32DiskDrive
            SelectWmiClassName = WmiClassNameWin32DiskDrive
        Case wmiEnumClassWin32DiskDrivePhysicalMedia
            SelectWmiClassName = WmiClassNameWin32DiskDrivePhysicalMedia
        Case wmiEnumClassWin32DiskDriveToDiskPartition
            SelectWmiClassName = WmiClassNameWin32DiskDriveToDiskPartition
        Case wmiEnumClassWin32DiskPartition
            SelectWmiClassName = WmiClassNameWin32DiskPartition
        Case wmiEnumClassWin32DiskQuota
            SelectWmiClassName = WmiClassNameWin32DiskQuota
        Case wmiEnumClassWin32DisplayConfiguration
            SelectWmiClassName = WmiClassNameWin32DisplayConfiguration
        Case wmiEnumClassWin32DisplayControllerConfiguration
            SelectWmiClassName = WmiClassNameWin32DisplayControllerConfiguration
        Case wmiEnumClassWin32DMAChannel
            SelectWmiClassName = WmiClassNameWin32DMAChannel
        Case wmiEnumClassWin32DriverForDevice
            SelectWmiClassName = WmiClassNameWin32DriverForDevice
        Case wmiEnumClassWin32DuplicateFileAction
            SelectWmiClassName = WmiClassNameWin32DuplicateFileAction
        Case wmiEnumClassWin32Environment
            SelectWmiClassName = WmiClassNameWin32Environment
        Case wmiEnumClassWin32EnvironmentSpecification
            SelectWmiClassName = WmiClassNameWin32EnvironmentSpecification
        Case wmiEnumClassWin32ExtensionInfoAction
            SelectWmiClassName = WmiClassNameWin32ExtensionInfoAction
        Case wmiEnumClassWin32Fan
            SelectWmiClassName = WmiClassNameWin32Fan
        Case wmiEnumClassWin32FileSpecification
            SelectWmiClassName = WmiClassNameWin32FileSpecification
        Case wmiEnumClassWin32FloppyController
            SelectWmiClassName = WmiClassNameWin32FloppyController
        Case wmiEnumClassWin32FloppyDrive
            SelectWmiClassName = WmiClassNameWin32FloppyDrive
        Case wmiEnumClassWin32FolderRedirection
            SelectWmiClassName = WmiClassNameWin32FolderRedirection
        Case wmiEnumClassWin32FolderRedirectionHealth
            SelectWmiClassName = WmiClassNameWin32FolderRedirectionHealth
        Case wmiEnumClassWin32FolderRedirectionHealthConfiguration
            SelectWmiClassName = WmiClassNameWin32FolderRedirectionHealthConfiguration
        Case wmiEnumClassWin32FolderRedirectionUserConfiguration
            SelectWmiClassName = WmiClassNameWin32FolderRedirectionUserConfiguration
        Case wmiEnumClassWin32FontInfoAction
            SelectWmiClassName = WmiClassNameWin32FontInfoAction
        Case wmiEnumClassWin32Group
            SelectWmiClassName = WmiClassNameWin32Group
        Case wmiEnumClassWin32GroupInDomain
            SelectWmiClassName = WmiClassNameWin32GroupInDomain
        Case wmiEnumClassWin32GroupUser
            SelectWmiClassName = WmiClassNameWin32GroupUser
        Case wmiEnumClassWin32HeatPipe
            SelectWmiClassName = WmiClassNameWin32HeatPipe
        Case wmiEnumClassWin32IDEController
            SelectWmiClassName = WmiClassNameWin32IDEController
        Case wmiEnumClassWin32IDEControllerDevice
            SelectWmiClassName = WmiClassNameWin32IDEControllerDevice
        Case wmiEnumClassWin32ImplementedCategory
            SelectWmiClassName = WmiClassNameWin32ImplementedCategory
        Case wmiEnumClassWin32InfraredDevice
            SelectWmiClassName = WmiClassNameWin32InfraredDevice
        Case wmiEnumClassWin32IniFileSpecification
            SelectWmiClassName = WmiClassNameWin32IniFileSpecification
        Case wmiEnumClassWin32InstalledProgramFramework
            SelectWmiClassName = WmiClassNameWin32InstalledProgramFramework
        Case wmiEnumClassWin32InstalledSoftwareElement
            SelectWmiClassName = WmiClassNameWin32InstalledSoftwareElement
        Case wmiEnumClassWin32InstalledStoreProgram
            SelectWmiClassName = WmiClassNameWin32InstalledStoreProgram
        Case wmiEnumClassWin32InstalledWin32Program
            SelectWmiClassName = WmiClassNameWin32InstalledWin32Program
        Case wmiEnumClassWin32IP4PersistedRouteTable
            SelectWmiClassName = WmiClassNameWin32IP4PersistedRouteTable
        Case wmiEnumClassWin32IP4RouteTable
            SelectWmiClassName = WmiClassNameWin32IP4RouteTable
        Case wmiEnumClassWin32IP4RouteTableEvent
            SelectWmiClassName = WmiClassNameWin32IP4RouteTableEvent
        Case wmiEnumClassWin32IRQResource
            SelectWmiClassName = WmiClassNameWin32IRQResource
        Case wmiEnumClassWin32JobObjectStatus
            SelectWmiClassName = WmiClassNameWin32JobObjectStatus
        Case wmiEnumClassWin32Keyboard
            SelectWmiClassName = WmiClassNameWin32Keyboard
        Case wmiEnumClassWin32LaunchCondition
            SelectWmiClassName = WmiClassNameWin32LaunchCondition
        Case wmiEnumClassWin32LoadOrderGroup
            SelectWmiClassName = WmiClassNameWin32LoadOrderGroup
        Case wmiEnumClassWin32LoadOrderGroupServiceDependencies
            SelectWmiClassName = WmiClassNameWin32LoadOrderGroupServiceDependencies
        Case wmiEnumClassWin32LoadOrderGroupServiceMembers
            SelectWmiClassName = WmiClassNameWin32LoadOrderGroupServiceMembers
        Case wmiEnumClassWin32LocalTime
            SelectWmiClassName = WmiClassNameWin32LocalTime
        Case wmiEnumClassWin32LoggedOnUser
            SelectWmiClassName = WmiClassNameWin32LoggedOnUser
        Case wmiEnumClassWin32LogicalDisk
            SelectWmiClassName = WmiClassNameWin32LogicalDisk
        Case wmiEnumClassWin32LogicalDiskRootDirectory
            SelectWmiClassName = WmiClassNameWin32LogicalDiskRootDirectory
        Case wmiEnumClassWin32LogicalDiskToPartition
            SelectWmiClassName = WmiClassNameWin32LogicalDiskToPartition
        Case wmiEnumClassWin32LogicalFileAccess
            SelectWmiClassName = WmiClassNameWin32LogicalFileAccess
        Case wmiEnumClassWin32LogicalFileAuditing
            SelectWmiClassName = WmiClassNameWin32LogicalFileAuditing
        Case wmiEnumClassWin32LogicalFileGroup
            SelectWmiClassName = WmiClassNameWin32LogicalFileGroup
        Case wmiEnumClassWin32LogicalFileOwner
            SelectWmiClassName = WmiClassNameWin32LogicalFileOwner
        Case wmiEnumClassWin32LogicalFileSecuritySetting
            SelectWmiClassName = WmiClassNameWin32LogicalFileSecuritySetting
        Case wmiEnumClassWin32LogicalProgramGroup
            SelectWmiClassName = WmiClassNameWin32LogicalProgramGroup
        Case wmiEnumClassWin32LogicalProgramGroupDirectory
            SelectWmiClassName = WmiClassNameWin32LogicalProgramGroupDirectory
        Case wmiEnumClassWin32LogicalProgramGroupItem
            SelectWmiClassName = WmiClassNameWin32LogicalProgramGroupItem
        Case wmiEnumClassWin32LogicalProgramGroupItemDataFile
            SelectWmiClassName = WmiClassNameWin32LogicalProgramGroupItemDataFile
        Case wmiEnumClassWin32LogicalShareAccess
            SelectWmiClassName = WmiClassNameWin32LogicalShareAccess
        Case wmiEnumClassWin32LogicalShareAuditing
            SelectWmiClassName = WmiClassNameWin32LogicalShareAuditing
        Case wmiEnumClassWin32LogicalShareSecuritySetting
            SelectWmiClassName = WmiClassNameWin32LogicalShareSecuritySetting
        Case wmiEnumClassWin32LogonSession
            SelectWmiClassName = WmiClassNameWin32LogonSession
        Case wmiEnumClassWin32LogonSessionMappedDisk
            SelectWmiClassName = WmiClassNameWin32LogonSessionMappedDisk
        Case wmiEnumClassWin32LUID
            SelectWmiClassName = WmiClassNameWin32LUID
        Case wmiEnumClassWin32LUIDandAttributes
            SelectWmiClassName = WmiClassNameWin32LUIDandAttributes
        Case wmiEnumClassWin32ManagedSystemElementResource
            SelectWmiClassName = WmiClassNameWin32ManagedSystemElementResource
        Case wmiEnumClassWin32MappedLogicalDisk
            SelectWmiClassName = WmiClassNameWin32MappedLogicalDisk
        Case wmiEnumClassWin32MemoryArray
            SelectWmiClassName = WmiClassNameWin32MemoryArray
        Case wmiEnumClassWin32MemoryArrayLocation
            SelectWmiClassName = WmiClassNameWin32MemoryArrayLocation
        Case wmiEnumClassWin32MemoryDevice
            SelectWmiClassName = WmiClassNameWin32MemoryDevice
        Case wmiEnumClassWin32MemoryDeviceArray
            SelectWmiClassName = WmiClassNameWin32MemoryDeviceArray
        Case wmiEnumClassWin32MemoryDeviceLocation
            SelectWmiClassName = WmiClassNameWin32MemoryDeviceLocation
        Case wmiEnumClassWin32MethodParameterClass
            SelectWmiClassName = WmiClassNameWin32MethodParameterClass
        Case wmiEnumClassWin32MIMEInfoAction
            SelectWmiClassName = WmiClassNameWin32MIMEInfoAction
        Case wmiEnumClassWin32ModuleLoadTrace
            SelectWmiClassName = WmiClassNameWin32ModuleLoadTrace
        Case wmiEnumClassWin32ModuleTrace
            SelectWmiClassName = WmiClassNameWin32ModuleTrace
        Case wmiEnumClassWin32MotherboardDevice
            SelectWmiClassName = WmiClassNameWin32MotherboardDevice
        Case wmiEnumClassWin32MountPoint
            SelectWmiClassName = WmiClassNameWin32MountPoint
        Case wmiEnumClassWin32MoveFileAction
            SelectWmiClassName = WmiClassNameWin32MoveFileAction
        Case wmiEnumClassWin32MSIResource
            SelectWmiClassName = WmiClassNameWin32MSIResource
        Case wmiEnumClassWin32NamedJobObject
            SelectWmiClassName = WmiClassNameWin32NamedJobObject
        Case wmiEnumClassWin32NamedJobObjectActgInfo
            SelectWmiClassName = WmiClassNameWin32NamedJobObjectActgInfo
        Case wmiEnumClassWin32NamedJobObjectLimit
            SelectWmiClassName = WmiClassNameWin32NamedJobObjectLimit
        Case wmiEnumClassWin32NamedJobObjectLimitSetting
            SelectWmiClassName = WmiClassNameWin32NamedJobObjectLimitSetting
        Case wmiEnumClassWin32NamedJobObjectProcess
            SelectWmiClassName = WmiClassNameWin32NamedJobObjectProcess
        Case wmiEnumClassWin32NamedJobObjectSecLimit
            SelectWmiClassName = WmiClassNameWin32NamedJobObjectSecLimit
        Case wmiEnumClassWin32NamedJobObjectSecLimitSetting
            SelectWmiClassName = WmiClassNameWin32NamedJobObjectSecLimitSetting
        Case wmiEnumClassWin32NamedJobObjectStatistics
            SelectWmiClassName = WmiClassNameWin32NamedJobObjectStatistics
        Case wmiEnumClassWin32NetworkAdapter
            SelectWmiClassName = WmiClassNameWin32NetworkAdapter
        Case wmiEnumClassWin32NetworkAdapterConfiguration
            SelectWmiClassName = WmiClassNameWin32NetworkAdapterConfiguration
        Case wmiEnumClassWin32NetworkAdapterSetting
            SelectWmiClassName = WmiClassNameWin32NetworkAdapterSetting
        Case wmiEnumClassWin32NetworkClient
            SelectWmiClassName = WmiClassNameWin32NetworkClient
        Case wmiEnumClassWin32NetworkConnection
            SelectWmiClassName = WmiClassNameWin32NetworkConnection
        Case wmiEnumClassWin32NetworkLoginProfile
            SelectWmiClassName = WmiClassNameWin32NetworkLoginProfile
        Case wmiEnumClassWin32NetworkProtocol
            SelectWmiClassName = WmiClassNameWin32NetworkProtocol
        Case wmiEnumClassWin32NTDomain
            SelectWmiClassName = WmiClassNameWin32NTDomain
        Case wmiEnumClassWin32NTEventlogFile
            SelectWmiClassName = WmiClassNameWin32NTEventlogFile
        Case wmiEnumClassWin32NTLogEvent
            SelectWmiClassName = WmiClassNameWin32NTLogEvent
        Case wmiEnumClassWin32NTLogEventComputer
            SelectWmiClassName = WmiClassNameWin32NTLogEventComputer
        Case wmiEnumClassWin32NTLogEventLog
            SelectWmiClassName = WmiClassNameWin32NTLogEventLog
        Case wmiEnumClassWin32NTLogEventUser
            SelectWmiClassName = WmiClassNameWin32NTLogEventUser
        Case wmiEnumClassWin32ODBCAttribute
            SelectWmiClassName = WmiClassNameWin32ODBCAttribute
        Case wmiEnumClassWin32ODBCDataSourceAttribute
            SelectWmiClassName = WmiClassNameWin32ODBCDataSourceAttribute
        Case wmiEnumClassWin32ODBCDataSourceSpecification
            SelectWmiClassName = WmiClassNameWin32ODBCDataSourceSpecification
        Case wmiEnumClassWin32ODBCDriverAttribute
            SelectWmiClassName = WmiClassNameWin32ODBCDriverAttribute
        Case wmiEnumClassWin32ODBCDriverSoftwareElement
            SelectWmiClassName = WmiClassNameWin32ODBCDriverSoftwareElement
        Case wmiEnumClassWin32ODBCDriverSpecification
            SelectWmiClassName = WmiClassNameWin32ODBCDriverSpecification
        Case wmiEnumClassWin32ODBCSourceAttribute
            SelectWmiClassName = WmiClassNameWin32ODBCSourceAttribute
        Case wmiEnumClassWin32ODBCTranslatorSpecification
            SelectWmiClassName = WmiClassNameWin32ODBCTranslatorSpecification
        Case wmiEnumClassWin32OfflineFilesAssociatedItems
            SelectWmiClassName = WmiClassNameWin32OfflineFilesAssociatedItems
        Case wmiEnumClassWin32OfflineFilesBackgroundSync
            SelectWmiClassName = WmiClassNameWin32OfflineFilesBackgroundSync
        Case wmiEnumClassWin32OfflineFilesCache
            SelectWmiClassName = WmiClassNameWin32OfflineFilesCache
        Case wmiEnumClassWin32OfflineFilesChangeInfo
            SelectWmiClassName = WmiClassNameWin32OfflineFilesChangeInfo
        Case wmiEnumClassWin32OfflineFilesConnectionInfo
            SelectWmiClassName = WmiClassNameWin32OfflineFilesConnectionInfo
        Case wmiEnumClassWin32OfflineFilesDirtyInfo
            SelectWmiClassName = WmiClassNameWin32OfflineFilesDirtyInfo
        Case wmiEnumClassWin32OfflineFilesDiskSpaceLimit
            SelectWmiClassName = WmiClassNameWin32OfflineFilesDiskSpaceLimit
        Case wmiEnumClassWin32OfflineFilesFileSysInfo
            SelectWmiClassName = WmiClassNameWin32OfflineFilesFileSysInfo
        Case wmiEnumClassWin32OfflineFilesHealth
            SelectWmiClassName = WmiClassNameWin32OfflineFilesHealth
        Case wmiEnumClassWin32OfflineFilesItem
            SelectWmiClassName = WmiClassNameWin32OfflineFilesItem
        Case wmiEnumClassWin32OfflineFilesMachineConfiguration
            SelectWmiClassName = WmiClassNameWin32OfflineFilesMachineConfiguration
        Case wmiEnumClassWin32OfflineFilesPinInfo
            SelectWmiClassName = WmiClassNameWin32OfflineFilesPinInfo
        Case wmiEnumClassWin32OfflineFilesSuspendInfo
            SelectWmiClassName = WmiClassNameWin32OfflineFilesSuspendInfo
        Case wmiEnumClassWin32OfflineFilesUserConfiguration
            SelectWmiClassName = WmiClassNameWin32OfflineFilesUserConfiguration
        Case wmiEnumClassWin32OnBoardDevice
            SelectWmiClassName = WmiClassNameWin32OnBoardDevice
        Case wmiEnumClassWin32OperatingSystem
            SelectWmiClassName = WmiClassNameWin32OperatingSystem
        Case wmiEnumClassWin32OperatingSystemAutochkSetting
            SelectWmiClassName = WmiClassNameWin32OperatingSystemAutochkSetting
        Case wmiEnumClassWin32OperatingSystemQFE
            SelectWmiClassName = WmiClassNameWin32OperatingSystemQFE
        Case wmiEnumClassWin32OptionalFeature
            SelectWmiClassName = WmiClassNameWin32OptionalFeature
        Case wmiEnumClassWin32OSRecoveryConfiguration
            SelectWmiClassName = WmiClassNameWin32OSRecoveryConfiguration
        Case wmiEnumClassWin32PageFile
            SelectWmiClassName = WmiClassNameWin32PageFile
        Case wmiEnumClassWin32PageFileElementSetting
            SelectWmiClassName = WmiClassNameWin32PageFileElementSetting
        Case wmiEnumClassWin32PageFileSetting
            SelectWmiClassName = WmiClassNameWin32PageFileSetting
        Case wmiEnumClassWin32PageFileUsage
            SelectWmiClassName = WmiClassNameWin32PageFileUsage
        Case wmiEnumClassWin32ParallelPort
            SelectWmiClassName = WmiClassNameWin32ParallelPort
        Case wmiEnumClassWin32Patch
            SelectWmiClassName = WmiClassNameWin32Patch
        Case wmiEnumClassWin32PatchFile
            SelectWmiClassName = WmiClassNameWin32PatchFile
        Case wmiEnumClassWin32PatchPackage
            SelectWmiClassName = WmiClassNameWin32PatchPackage
        Case wmiEnumClassWin32PCMCIAController
            SelectWmiClassName = WmiClassNameWin32PCMCIAController
        Case wmiEnumClassWin32PhysicalMedia
            SelectWmiClassName = WmiClassNameWin32PhysicalMedia
        Case wmiEnumClassWin32PhysicalMemory
            SelectWmiClassName = WmiClassNameWin32PhysicalMemory
        Case wmiEnumClassWin32PhysicalMemoryArray
            SelectWmiClassName = WmiClassNameWin32PhysicalMemoryArray
        Case wmiEnumClassWin32PhysicalMemoryLocation
            SelectWmiClassName = WmiClassNameWin32PhysicalMemoryLocation
        Case wmiEnumClassWin32PingStatus
            SelectWmiClassName = WmiClassNameWin32PingStatus
        Case wmiEnumClassWin32PNPAllocatedResource
            SelectWmiClassName = WmiClassNameWin32PNPAllocatedResource
        Case wmiEnumClassWin32PnPDevice
            SelectWmiClassName = WmiClassNameWin32PnPDevice
        Case wmiEnumClassWin32PnPEntity
            SelectWmiClassName = WmiClassNameWin32PnPEntity
        Case wmiEnumClassWin32PnPSignedDriver
            SelectWmiClassName = WmiClassNameWin32PnPSignedDriver
        Case wmiEnumClassWin32PnPSignedDriverCIMDataFile
            SelectWmiClassName = WmiClassNameWin32PnPSignedDriverCIMDataFile
        Case wmiEnumClassWin32PointingDevice
            SelectWmiClassName = WmiClassNameWin32PointingDevice
        Case wmiEnumClassWin32PortableBattery
            SelectWmiClassName = WmiClassNameWin32PortableBattery
        Case wmiEnumClassWin32PortConnector
            SelectWmiClassName = WmiClassNameWin32PortConnector
        Case wmiEnumClassWin32PortResource
            SelectWmiClassName = WmiClassNameWin32PortResource
        Case wmiEnumClassWin32POTSModem
            SelectWmiClassName = WmiClassNameWin32POTSModem
        Case wmiEnumClassWin32POTSModemToSerialPort
            SelectWmiClassName = WmiClassNameWin32POTSModemToSerialPort
        Case wmiEnumClassWin32PowerManagementEvent
            SelectWmiClassName = WmiClassNameWin32PowerManagementEvent
        Case wmiEnumClassWin32Printer
            SelectWmiClassName = WmiClassNameWin32Printer
        Case wmiEnumClassWin32PrinterConfiguration
            SelectWmiClassName = WmiClassNameWin32PrinterConfiguration
        Case wmiEnumClassWin32PrinterController
            SelectWmiClassName = WmiClassNameWin32PrinterController
        Case wmiEnumClassWin32PrinterDriver
            SelectWmiClassName = WmiClassNameWin32PrinterDriver
        Case wmiEnumClassWin32PrinterDriverDll
            SelectWmiClassName = WmiClassNameWin32PrinterDriverDll
        Case wmiEnumClassWin32PrinterSetting
            SelectWmiClassName = WmiClassNameWin32PrinterSetting
        Case wmiEnumClassWin32PrinterShare
            SelectWmiClassName = WmiClassNameWin32PrinterShare
        Case wmiEnumClassWin32PrintJob
            SelectWmiClassName = WmiClassNameWin32PrintJob
        Case wmiEnumClassWin32PrivilegesStatus
            SelectWmiClassName = WmiClassNameWin32PrivilegesStatus
        Case wmiEnumClassWin32Process
            SelectWmiClassName = WmiClassNameWin32Process
        Case wmiEnumClassWin32Processor
            SelectWmiClassName = WmiClassNameWin32Processor
        Case wmiEnumClassWin32ProcessStartTrace
            SelectWmiClassName = WmiClassNameWin32ProcessStartTrace
        Case wmiEnumClassWin32ProcessStartup
            SelectWmiClassName = WmiClassNameWin32ProcessStartup
        Case wmiEnumClassWin32ProcessStopTrace
            SelectWmiClassName = WmiClassNameWin32ProcessStopTrace
        Case wmiEnumClassWin32ProcessTrace
            SelectWmiClassName = WmiClassNameWin32ProcessTrace
        Case wmiEnumClassWin32Product
            SelectWmiClassName = WmiClassNameWin32Product
        Case wmiEnumClassWin32ProductCheck
            SelectWmiClassName = WmiClassNameWin32ProductCheck
        Case wmiEnumClassWin32ProductResource
            SelectWmiClassName = WmiClassNameWin32ProductResource
        Case wmiEnumClassWin32ProductSoftwareFeatures
            SelectWmiClassName = WmiClassNameWin32ProductSoftwareFeatures
        Case wmiEnumClassWin32ProgIDSpecification
            SelectWmiClassName = WmiClassNameWin32ProgIDSpecification
        Case wmiEnumClassWin32ProgramGroupContents
            SelectWmiClassName = WmiClassNameWin32ProgramGroupContents
        Case wmiEnumClassWin32ProgramGroupOrItem
            SelectWmiClassName = WmiClassNameWin32ProgramGroupOrItem
        Case wmiEnumClassWin32Property
            SelectWmiClassName = WmiClassNameWin32Property
        Case wmiEnumClassWin32ProtocolBinding
            SelectWmiClassName = WmiClassNameWin32ProtocolBinding
        Case wmiEnumClassWin32PublishComponentAction
            SelectWmiClassName = WmiClassNameWin32PublishComponentAction
        Case wmiEnumClassWin32QuickFixEngineering
            SelectWmiClassName = WmiClassNameWin32QuickFixEngineering
        Case wmiEnumClassWin32QuotaSetting
            SelectWmiClassName = WmiClassNameWin32QuotaSetting
        Case wmiEnumClassWin32Refrigeration
            SelectWmiClassName = WmiClassNameWin32Refrigeration
        Case wmiEnumClassWin32Registry
            SelectWmiClassName = WmiClassNameWin32Registry
        Case wmiEnumClassWin32RegistryAction
            SelectWmiClassName = WmiClassNameWin32RegistryAction
        Case wmiEnumClassWin32Reliability
            SelectWmiClassName = WmiClassNameWin32Reliability
        Case wmiEnumClassWin32ReliabilityRecords
            SelectWmiClassName = WmiClassNameWin32ReliabilityRecords
        Case wmiEnumClassWin32ReliabilityStabilityMetrics
            SelectWmiClassName = WmiClassNameWin32ReliabilityStabilityMetrics
        Case wmiEnumClassWin32RemoveFileAction
            SelectWmiClassName = WmiClassNameWin32RemoveFileAction
        Case wmiEnumClassWin32RemoveIniAction
            SelectWmiClassName = WmiClassNameWin32RemoveIniAction
        Case wmiEnumClassWin32ReserveCost
            SelectWmiClassName = WmiClassNameWin32ReserveCost
        Case wmiEnumClassWin32RoamingProfileBackgroundUploadParams
            SelectWmiClassName = WmiClassNameWin32RoamingProfileBackgroundUploadParams
        Case wmiEnumClassWin32RoamingProfileMachineConfiguration
            SelectWmiClassName = WmiClassNameWin32RoamingProfileMachineConfiguration
        Case wmiEnumClassWin32RoamingProfileSlowLinkParams
            SelectWmiClassName = WmiClassNameWin32RoamingProfileSlowLinkParams
        Case wmiEnumClassWin32RoamingProfileUserConfiguration
            SelectWmiClassName = WmiClassNameWin32RoamingProfileUserConfiguration
        Case wmiEnumClassWin32RoamingUserHealthConfiguration
            SelectWmiClassName = WmiClassNameWin32RoamingUserHealthConfiguration
        Case wmiEnumClassWin32ScheduledJob
            SelectWmiClassName = WmiClassNameWin32ScheduledJob
        Case wmiEnumClassWin32SCSIController
            SelectWmiClassName = WmiClassNameWin32SCSIController
        Case wmiEnumClassWin32SCSIControllerDevice
            SelectWmiClassName = WmiClassNameWin32SCSIControllerDevice
        Case wmiEnumClassWin32SecurityDescriptor
            SelectWmiClassName = WmiClassNameWin32SecurityDescriptor
        Case wmiEnumClassWin32SecurityDescriptorHelper
            SelectWmiClassName = WmiClassNameWin32SecurityDescriptorHelper
        Case wmiEnumClassWin32SecuritySetting
            SelectWmiClassName = WmiClassNameWin32SecuritySetting
        Case wmiEnumClassWin32SecuritySettingAccess
            SelectWmiClassName = WmiClassNameWin32SecuritySettingAccess
        Case wmiEnumClassWin32SecuritySettingAuditing
            SelectWmiClassName = WmiClassNameWin32SecuritySettingAuditing
        Case wmiEnumClassWin32SecuritySettingGroup
            SelectWmiClassName = WmiClassNameWin32SecuritySettingGroup
        Case wmiEnumClassWin32SecuritySettingOfLogicalFile
            SelectWmiClassName = WmiClassNameWin32SecuritySettingOfLogicalFile
        Case wmiEnumClassWin32SecuritySettingOfLogicalShare
            SelectWmiClassName = WmiClassNameWin32SecuritySettingOfLogicalShare
        Case wmiEnumClassWin32SecuritySettingOfObject
            SelectWmiClassName = WmiClassNameWin32SecuritySettingOfObject
        Case wmiEnumClassWin32SecuritySettingOwner
            SelectWmiClassName = WmiClassNameWin32SecuritySettingOwner
        Case wmiEnumClassWin32SelfRegModuleAction
            SelectWmiClassName = WmiClassNameWin32SelfRegModuleAction
        Case wmiEnumClassWin32SerialPort
            SelectWmiClassName = WmiClassNameWin32SerialPort
        Case wmiEnumClassWin32SerialPortConfiguration
            SelectWmiClassName = WmiClassNameWin32SerialPortConfiguration
        Case wmiEnumClassWin32SerialPortSetting
            SelectWmiClassName = WmiClassNameWin32SerialPortSetting
        Case wmiEnumClassWin32ServerConnection
            SelectWmiClassName = WmiClassNameWin32ServerConnection
        Case wmiEnumClassWin32ServerFeature
            SelectWmiClassName = WmiClassNameWin32ServerFeature
        Case wmiEnumClassWin32ServerSession
            SelectWmiClassName = WmiClassNameWin32ServerSession
        Case wmiEnumClassWin32Service
            SelectWmiClassName = WmiClassNameWin32Service
        Case wmiEnumClassWin32ServiceControl
            SelectWmiClassName = WmiClassNameWin32ServiceControl
        Case wmiEnumClassWin32ServiceSpecification
            SelectWmiClassName = WmiClassNameWin32ServiceSpecification
        Case wmiEnumClassWin32ServiceSpecificationService
            SelectWmiClassName = WmiClassNameWin32ServiceSpecificationService
        Case wmiEnumClassWin32Session
            SelectWmiClassName = WmiClassNameWin32Session
        Case wmiEnumClassWin32SessionConnection
            SelectWmiClassName = WmiClassNameWin32SessionConnection
        Case wmiEnumClassWin32SessionProcess
            SelectWmiClassName = WmiClassNameWin32SessionProcess
        Case wmiEnumClassWin32SessionResource
            SelectWmiClassName = WmiClassNameWin32SessionResource
        Case wmiEnumClassWin32SettingCheck
            SelectWmiClassName = WmiClassNameWin32SettingCheck
        Case wmiEnumClassWin32ShadowBy
            SelectWmiClassName = WmiClassNameWin32ShadowBy
        Case wmiEnumClassWin32ShadowContext
            SelectWmiClassName = WmiClassNameWin32ShadowContext
        Case wmiEnumClassWin32ShadowCopy
            SelectWmiClassName = WmiClassNameWin32ShadowCopy
        Case wmiEnumClassWin32ShadowDiffVolumeSupport
            SelectWmiClassName = WmiClassNameWin32ShadowDiffVolumeSupport
        Case wmiEnumClassWin32ShadowFor
            SelectWmiClassName = WmiClassNameWin32ShadowFor
        Case wmiEnumClassWin32ShadowOn
            SelectWmiClassName = WmiClassNameWin32ShadowOn
        Case wmiEnumClassWin32ShadowProvider
            SelectWmiClassName = WmiClassNameWin32ShadowProvider
        Case wmiEnumClassWin32ShadowStorage
            SelectWmiClassName = WmiClassNameWin32ShadowStorage
        Case wmiEnumClassWin32ShadowVolumeSupport
            SelectWmiClassName = WmiClassNameWin32ShadowVolumeSupport
        Case wmiEnumClassWin32Share
            SelectWmiClassName = WmiClassNameWin32Share
        Case wmiEnumClassWin32ShareToDirectory
            SelectWmiClassName = WmiClassNameWin32ShareToDirectory
        Case wmiEnumClassWin32ShortcutAction
            SelectWmiClassName = WmiClassNameWin32ShortcutAction
        Case wmiEnumClassWin32ShortcutFile
            SelectWmiClassName = WmiClassNameWin32ShortcutFile
        Case wmiEnumClassWin32ShortcutSAP
            SelectWmiClassName = WmiClassNameWin32ShortcutSAP
        Case wmiEnumClassWin32SID
            SelectWmiClassName = WmiClassNameWin32SID
        Case wmiEnumClassWin32SIDandAttributes
            SelectWmiClassName = WmiClassNameWin32SIDandAttributes
        Case wmiEnumClassWin32SMBIOSMemory
            SelectWmiClassName = WmiClassNameWin32SMBIOSMemory
        Case wmiEnumClassWin32SoftwareElement
            SelectWmiClassName = WmiClassNameWin32SoftwareElement
        Case wmiEnumClassWin32SoftwareElementAction
            SelectWmiClassName = WmiClassNameWin32SoftwareElementAction
        Case wmiEnumClassWin32SoftwareElementCheck
            SelectWmiClassName = WmiClassNameWin32SoftwareElementCheck
        Case wmiEnumClassWin32SoftwareElementCondition
            SelectWmiClassName = WmiClassNameWin32SoftwareElementCondition
        Case wmiEnumClassWin32SoftwareElementResource
            SelectWmiClassName = WmiClassNameWin32SoftwareElementResource
        Case wmiEnumClassWin32SoftwareFeature
            SelectWmiClassName = WmiClassNameWin32SoftwareFeature
        Case wmiEnumClassWin32SoftwareFeatureAction
            SelectWmiClassName = WmiClassNameWin32SoftwareFeatureAction
        Case wmiEnumClassWin32SoftwareFeatureCheck
            SelectWmiClassName = WmiClassNameWin32SoftwareFeatureCheck
        Case wmiEnumClassWin32SoftwareFeatureParent
            SelectWmiClassName = WmiClassNameWin32SoftwareFeatureParent
        Case wmiEnumClassWin32SoftwareFeatureSoftwareElements
            SelectWmiClassName = WmiClassNameWin32SoftwareFeatureSoftwareElements
        Case wmiEnumClassWin32SoundDevice
            SelectWmiClassName = WmiClassNameWin32SoundDevice
        Case wmiEnumClassWin32StartupCommand
            SelectWmiClassName = WmiClassNameWin32StartupCommand
        Case wmiEnumClassWin32SubDirectory
            SelectWmiClassName = WmiClassNameWin32SubDirectory
        Case wmiEnumClassWin32SubSession
            SelectWmiClassName = WmiClassNameWin32SubSession
        Case wmiEnumClassWin32SystemAccount
            SelectWmiClassName = WmiClassNameWin32SystemAccount
        Case wmiEnumClassWin32SystemBIOS
            SelectWmiClassName = WmiClassNameWin32SystemBIOS
        Case wmiEnumClassWin32SystemBootConfiguration
            SelectWmiClassName = WmiClassNameWin32SystemBootConfiguration
        Case wmiEnumClassWin32SystemConfigurationChangeEvent
            SelectWmiClassName = WmiClassNameWin32SystemConfigurationChangeEvent
        Case wmiEnumClassWin32SystemDesktop
            SelectWmiClassName = WmiClassNameWin32SystemDesktop
        Case wmiEnumClassWin32SystemDevices
            SelectWmiClassName = WmiClassNameWin32SystemDevices
        Case wmiEnumClassWin32SystemDriver
            SelectWmiClassName = WmiClassNameWin32SystemDriver
        Case wmiEnumClassWin32SystemDriverPNPEntity
            SelectWmiClassName = WmiClassNameWin32SystemDriverPNPEntity
        Case wmiEnumClassWin32SystemEnclosure
            SelectWmiClassName = WmiClassNameWin32SystemEnclosure
        Case wmiEnumClassWin32SystemLoadOrderGroups
            SelectWmiClassName = WmiClassNameWin32SystemLoadOrderGroups
        Case wmiEnumClassWin32SystemMemoryResource
            SelectWmiClassName = WmiClassNameWin32SystemMemoryResource
        Case wmiEnumClassWin32SystemNetworkConnections
            SelectWmiClassName = WmiClassNameWin32SystemNetworkConnections
        Case wmiEnumClassWin32SystemOperatingSystem
            SelectWmiClassName = WmiClassNameWin32SystemOperatingSystem
        Case wmiEnumClassWin32SystemPartitions
            SelectWmiClassName = WmiClassNameWin32SystemPartitions
        Case wmiEnumClassWin32SystemProcesses
            SelectWmiClassName = WmiClassNameWin32SystemProcesses
        Case wmiEnumClassWin32SystemProgramGroups
            SelectWmiClassName = WmiClassNameWin32SystemProgramGroups
        Case wmiEnumClassWin32SystemResources
            SelectWmiClassName = WmiClassNameWin32SystemResources
        Case wmiEnumClassWin32SystemServices
            SelectWmiClassName = WmiClassNameWin32SystemServices
        Case wmiEnumClassWin32SystemSetting
            SelectWmiClassName = WmiClassNameWin32SystemSetting
        Case wmiEnumClassWin32SystemSlot
            SelectWmiClassName = WmiClassNameWin32SystemSlot
        Case wmiEnumClassWin32SystemSystemDriver
            SelectWmiClassName = WmiClassNameWin32SystemSystemDriver
        Case wmiEnumClassWin32SystemTimeZone
            SelectWmiClassName = WmiClassNameWin32SystemTimeZone
        Case wmiEnumClassWin32SystemTrace
            SelectWmiClassName = WmiClassNameWin32SystemTrace
        Case wmiEnumClassWin32SystemUsers
            SelectWmiClassName = WmiClassNameWin32SystemUsers
        Case wmiEnumClassWin32TapeDrive
            SelectWmiClassName = WmiClassNameWin32TapeDrive
        Case wmiEnumClassWin32TCPIPPrinterPort
            SelectWmiClassName = WmiClassNameWin32TCPIPPrinterPort
        Case wmiEnumClassWin32TemperatureProbe
            SelectWmiClassName = WmiClassNameWin32TemperatureProbe
        Case wmiEnumClassWin32TerminalService
            SelectWmiClassName = WmiClassNameWin32TerminalService
        Case wmiEnumClassWin32Thread
            SelectWmiClassName = WmiClassNameWin32Thread
        Case wmiEnumClassWin32ThreadStartTrace
            SelectWmiClassName = WmiClassNameWin32ThreadStartTrace
        Case wmiEnumClassWin32ThreadStopTrace
            SelectWmiClassName = WmiClassNameWin32ThreadStopTrace
        Case wmiEnumClassWin32ThreadTrace
            SelectWmiClassName = WmiClassNameWin32ThreadTrace
        Case wmiEnumClassWin32TimeZone
            SelectWmiClassName = WmiClassNameWin32TimeZone
        Case wmiEnumClassWin32TokenGroups
            SelectWmiClassName = WmiClassNameWin32TokenGroups
        Case wmiEnumClassWin32TokenPrivileges
            SelectWmiClassName = WmiClassNameWin32TokenPrivileges
        Case wmiEnumClassWin32Trustee
            SelectWmiClassName = WmiClassNameWin32Trustee
        Case wmiEnumClassWin32TypeLibraryAction
            SelectWmiClassName = WmiClassNameWin32TypeLibraryAction
        Case wmiEnumClassWin32USBController
            SelectWmiClassName = WmiClassNameWin32USBController
        Case wmiEnumClassWin32USBControllerDevice
            SelectWmiClassName = WmiClassNameWin32USBControllerDevice
        Case wmiEnumClassWin32USBHub
            SelectWmiClassName = WmiClassNameWin32USBHub
        Case wmiEnumClassWin32UserAccount
            SelectWmiClassName = WmiClassNameWin32UserAccount
        Case wmiEnumClassWin32UserDesktop
            SelectWmiClassName = WmiClassNameWin32UserDesktop
        Case wmiEnumClassWin32UserInDomain
            SelectWmiClassName = WmiClassNameWin32UserInDomain
        Case wmiEnumClassWin32UserProfile
            SelectWmiClassName = WmiClassNameWin32UserProfile
        Case wmiEnumClassWin32UserStateConfigurationControls
            SelectWmiClassName = WmiClassNameWin32UserStateConfigurationControls
        Case wmiEnumClassWin32UTCTime
            SelectWmiClassName = WmiClassNameWin32UTCTime
        Case wmiEnumClassWin32VideoConfiguration
            SelectWmiClassName = WmiClassNameWin32VideoConfiguration
        Case wmiEnumClassWin32VideoController
            SelectWmiClassName = WmiClassNameWin32VideoController
        Case wmiEnumClassWin32VideoSettings
            SelectWmiClassName = WmiClassNameWin32VideoSettings
        Case wmiEnumClassWin32VoltageProbe
            SelectWmiClassName = WmiClassNameWin32VoltageProbe
        Case wmiEnumClassWin32Volume
            SelectWmiClassName = WmiClassNameWin32Volume
        Case wmiEnumClassWin32VolumeChangeEvent
            SelectWmiClassName = WmiClassNameWin32VolumeChangeEvent
        Case wmiEnumClassWin32VolumeQuota
            SelectWmiClassName = WmiClassNameWin32VolumeQuota
        Case wmiEnumClassWin32VolumeQuotaSetting
            SelectWmiClassName = WmiClassNameWin32VolumeQuotaSetting
        Case wmiEnumClassWin32VolumeUserQuota
            SelectWmiClassName = WmiClassNameWin32VolumeUserQuota
        Case wmiEnumClassWin32WinSAT
            SelectWmiClassName = WmiClassNameWin32WinSAT
        Case wmiEnumClassWin32WMIElementSetting
            SelectWmiClassName = WmiClassNameWin32WMIElementSetting
        Case wmiEnumClassWin32WMISetting
            SelectWmiClassName = WmiClassNameWin32WMISetting
        Case wmiEnumClassWin32Perf
            SelectWmiClassName = WmiClassNameWin32Perf
        Case wmiEnumClassWin32PerfFormattedData
            SelectWmiClassName = WmiClassNameWin32PerfFormattedData
        Case wmiEnumClassWin32PerfFormattedDataAFDCountersMicrosoftWinsockBSP
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataAFDCountersMicrosoftWinsockBSP
        Case wmiEnumClassWin32PerfFormattedDataAPPPOOLCountersProviderAPPPOOLWAS
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataAPPPOOLCountersProviderAPPPOOLWAS
        Case wmiEnumClassWin32PerfFormattedDataASPActiveServerPages
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataASPActiveServerPages
        Case wmiEnumClassWin32PerfFormattedDataASPNETASPNET
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataASPNETASPNET
        Case wmiEnumClassWin32PerfFormattedDataASPNETASPNETApplications
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataASPNETASPNETApplications
        Case wmiEnumClassWin32PerfFormattedDataASPNET2050727ASPNETAppsv2050727
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataASPNET2050727ASPNETAppsv2050727
        Case wmiEnumClassWin32PerfFormattedDataASPNET2050727ASPNETv2050727
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataASPNET2050727ASPNETv2050727
        Case wmiEnumClassWin32PerfFormattedDataASPNET4030319ASPNETAppsv4030319
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataASPNET4030319ASPNETAppsv4030319
        Case wmiEnumClassWin32PerfFormattedDataASPNET4030319ASPNETv4030319
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataASPNET4030319ASPNETv4030319
        Case wmiEnumClassWin32PerfFormattedDataaspnetstateASPNETStateService
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataaspnetstateASPNETStateService
        Case wmiEnumClassWin32PerfFormattedDataAuthorizationManagerAuthorizationManagerApplications
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataAuthorizationManagerAuthorizationManagerApplications
        Case wmiEnumClassWin32PerfFormattedDataBalancerStatsHyperVDynamicMemoryBalancer
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataBalancerStatsHyperVDynamicMemoryBalancer
        Case wmiEnumClassWin32PerfFormattedDataBalancerStatsHyperVDynamicMemoryVM
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataBalancerStatsHyperVDynamicMemoryVM
        Case wmiEnumClassWin32PerfFormattedDataBITSBITSNetUtilization
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataBITSBITSNetUtilization
        Case wmiEnumClassWin32PerfFormattedDataCountersDNS64Global
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersDNS64Global
        Case wmiEnumClassWin32PerfFormattedDataCountersEventTracingforWindows
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersEventTracingforWindows
        Case wmiEnumClassWin32PerfFormattedDataCountersEventTracingforWindowsSession
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersEventTracingforWindowsSession
        Case wmiEnumClassWin32PerfFormattedDataCountersFileSystemDiskActivity
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersFileSystemDiskActivity
        Case wmiEnumClassWin32PerfFormattedDataCountersGenericIKEv1AuthIPandIKEv2
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersGenericIKEv1AuthIPandIKEv2
        Case wmiEnumClassWin32PerfFormattedDataCountersHTTPService
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersHTTPService
        Case wmiEnumClassWin32PerfFormattedDataCountersHTTPServiceRequestQueues
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersHTTPServiceRequestQueues
        Case wmiEnumClassWin32PerfFormattedDataCountersHTTPServiceUrlGroups
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersHTTPServiceUrlGroups
        Case wmiEnumClassWin32PerfFormattedDataCountersHyperVDynamicMemoryIntegrationService
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersHyperVDynamicMemoryIntegrationService
        Case wmiEnumClassWin32PerfFormattedDataCountersHyperVVirtualMachineBusPipes
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersHyperVVirtualMachineBusPipes
        Case wmiEnumClassWin32PerfFormattedDataCountersIPHTTPSGlobal
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersIPHTTPSGlobal
        Case wmiEnumClassWin32PerfFormattedDataCountersIPHTTPSSession
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersIPHTTPSSession
        Case wmiEnumClassWin32PerfFormattedDataCountersIPsecAuthIPIPv4
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersIPsecAuthIPIPv4
        Case wmiEnumClassWin32PerfFormattedDataCountersIPsecAuthIPIPv6
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersIPsecAuthIPIPv6
        Case wmiEnumClassWin32PerfFormattedDataCountersIPsecConnections
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersIPsecConnections
        Case wmiEnumClassWin32PerfFormattedDataCountersIPsecDoSProtection
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersIPsecDoSProtection
        Case wmiEnumClassWin32PerfFormattedDataCountersIPsecDriver
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersIPsecDriver
        Case wmiEnumClassWin32PerfFormattedDataCountersIPsecIKEv1IPv4
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersIPsecIKEv1IPv4
        Case wmiEnumClassWin32PerfFormattedDataCountersIPsecIKEv1IPv6
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersIPsecIKEv1IPv6
        Case wmiEnumClassWin32PerfFormattedDataCountersIPsecIKEv2IPv4
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersIPsecIKEv2IPv4
        Case wmiEnumClassWin32PerfFormattedDataCountersIPsecIKEv2IPv6
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersIPsecIKEv2IPv6
        Case wmiEnumClassWin32PerfFormattedDataCountersNetlogon
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersNetlogon
        Case wmiEnumClassWin32PerfFormattedDataCountersNetworkQoSPolicy
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersNetworkQoSPolicy
        Case wmiEnumClassWin32PerfFormattedDataCountersPacerFlow
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersPacerFlow
        Case wmiEnumClassWin32PerfFormattedDataCountersPacerPipe
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersPacerPipe
        Case wmiEnumClassWin32PerfFormattedDataCountersPacketDirectECUtilization
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersPacketDirectECUtilization
        Case wmiEnumClassWin32PerfFormattedDataCountersPacketDirectQueueDepth
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersPacketDirectQueueDepth
        Case wmiEnumClassWin32PerfFormattedDataCountersPacketDirectReceiveCounters
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersPacketDirectReceiveCounters
        Case wmiEnumClassWin32PerfFormattedDataCountersPacketDirectReceiveFilters
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersPacketDirectReceiveFilters
        Case wmiEnumClassWin32PerfFormattedDataCountersPacketDirectTransmitCounters
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersPacketDirectTransmitCounters
        Case wmiEnumClassWin32PerfFormattedDataCountersPerProcessorNetworkActivityCycles
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersPerProcessorNetworkActivityCycles
        Case wmiEnumClassWin32PerfFormattedDataCountersPerProcessorNetworkInterfaceCardActivity
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersPerProcessorNetworkInterfaceCardActivity
        Case wmiEnumClassWin32PerfFormattedDataCountersPhysicalNetworkInterfaceCardActivity
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersPhysicalNetworkInterfaceCardActivity
        Case wmiEnumClassWin32PerfFormattedDataCountersPowerShellWorkflow
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersPowerShellWorkflow
        Case wmiEnumClassWin32PerfFormattedDataCountersProcessorInformation
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersProcessorInformation
        Case wmiEnumClassWin32PerfFormattedDataCountersRDMAActivity
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersRDMAActivity
        Case wmiEnumClassWin32PerfFormattedDataCountersRemoteFXGraphics
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersRemoteFXGraphics
        Case wmiEnumClassWin32PerfFormattedDataCountersRemoteFXNetwork
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersRemoteFXNetwork
        Case wmiEnumClassWin32PerfFormattedDataCountersSMBClientShares
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersSMBClientShares
        Case wmiEnumClassWin32PerfFormattedDataCountersSMBServer
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersSMBServer
        Case wmiEnumClassWin32PerfFormattedDataCountersSMBServerSessions
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersSMBServerSessions
        Case wmiEnumClassWin32PerfFormattedDataCountersSMBServerShares
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersSMBServerShares
        Case wmiEnumClassWin32PerfFormattedDataCountersStorageSpacesTier
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersStorageSpacesTier
        Case wmiEnumClassWin32PerfFormattedDataCountersStorageSpacesWriteCache
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersStorageSpacesWriteCache
        Case wmiEnumClassWin32PerfFormattedDataCountersSynchronization
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersSynchronization
        Case wmiEnumClassWin32PerfFormattedDataCountersSynchronizationNuma
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersSynchronizationNuma
        Case wmiEnumClassWin32PerfFormattedDataCountersTeredoClient
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersTeredoClient
        Case wmiEnumClassWin32PerfFormattedDataCountersTeredoRelay
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersTeredoRelay
        Case wmiEnumClassWin32PerfFormattedDataCountersTeredoServer
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersTeredoServer
        Case wmiEnumClassWin32PerfFormattedDataCountersThermalZoneInformation
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersThermalZoneInformation
        Case wmiEnumClassWin32PerfFormattedDataCountersWFP
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersWFP
        Case wmiEnumClassWin32PerfFormattedDataCountersWFPv4
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersWFPv4
        Case wmiEnumClassWin32PerfFormattedDataCountersWFPv6
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersWFPv6
        Case wmiEnumClassWin32PerfFormattedDataCountersWSManQuotaStatistics
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersWSManQuotaStatistics
        Case wmiEnumClassWin32PerfFormattedDataCountersXHCICommonBuffer
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersXHCICommonBuffer
        Case wmiEnumClassWin32PerfFormattedDataCountersXHCIInterrupter
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersXHCIInterrupter
        Case wmiEnumClassWin32PerfFormattedDataCountersXHCITransferRing
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataCountersXHCITransferRing
        Case wmiEnumClassWin32PerfFormattedDataDdmCounterProviderRAS
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataDdmCounterProviderRAS
        Case wmiEnumClassWin32PerfFormattedDataDeliveryOptimizationDeliveryOptimizationSwarm
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataDeliveryOptimizationDeliveryOptimizationSwarm
        Case wmiEnumClassWin32PerfFormattedDataDistributedRoutingTablePerfDistributedRoutingTable
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataDistributedRoutingTablePerfDistributedRoutingTable
        Case wmiEnumClassWin32PerfFormattedDataESENTDatabase
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataESENTDatabase
        Case wmiEnumClassWin32PerfFormattedDataESENTDatabaseInstances
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataESENTDatabaseInstances
        Case wmiEnumClassWin32PerfFormattedDataESENTDatabaseTableClasses
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataESENTDatabaseTableClasses
        Case wmiEnumClassWin32PerfFormattedDataEthernetPerfProviderHyperVLegacyNetworkAdapter
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataEthernetPerfProviderHyperVLegacyNetworkAdapter
        Case wmiEnumClassWin32PerfFormattedDataFaxServiceFaxService
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataFaxServiceFaxService
        Case wmiEnumClassWin32PerfFormattedDataftpsvcMicrosoftFTPService
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataftpsvcMicrosoftFTPService
        Case wmiEnumClassWin32PerfFormattedDataGmoPerfProviderHyperVVMSaveSnapshotandRestore
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataGmoPerfProviderHyperVVMSaveSnapshotandRestore
        Case wmiEnumClassWin32PerfFormattedDataHvStatsHyperVHypervisor
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataHvStatsHyperVHypervisor
        Case wmiEnumClassWin32PerfFormattedDataHvStatsHyperVHypervisorLogicalProcessor
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataHvStatsHyperVHypervisorLogicalProcessor
        Case wmiEnumClassWin32PerfFormattedDataHvStatsHyperVHypervisorPartition
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataHvStatsHyperVHypervisorPartition
        Case wmiEnumClassWin32PerfFormattedDataHvStatsHyperVHypervisorRootPartition
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataHvStatsHyperVHypervisorRootPartition
        Case wmiEnumClassWin32PerfFormattedDataHvStatsHyperVHypervisorRootVirtualProcessor
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataHvStatsHyperVHypervisorRootVirtualProcessor
        Case wmiEnumClassWin32PerfFormattedDataHvStatsHyperVHypervisorVirtualProcessor
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataHvStatsHyperVHypervisorVirtualProcessor
        Case wmiEnumClassWin32PerfFormattedDataIdePerfProviderHyperVVirtualIDEController
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataIdePerfProviderHyperVVirtualIDEController
        Case wmiEnumClassWin32PerfFormattedDataLocalSessionManagerTerminalServices
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataLocalSessionManagerTerminalServices
        Case wmiEnumClassWin32PerfFormattedDataLsaSecurityPerProcessStatistics
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataLsaSecurityPerProcessStatistics
        Case wmiEnumClassWin32PerfFormattedDataLsaSecuritySystemWideStatistics
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataLsaSecuritySystemWideStatistics
        Case wmiEnumClassWin32PerfFormattedDataMicrosoftWindowsBitLockerDriverCountersProviderBitLocker
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataMicrosoftWindowsBitLockerDriverCountersProviderBitLocker
        Case wmiEnumClassWin32PerfFormattedDataMicrosoftWindowsRemoteDesktopServicesRemoteFXSynth3dvscRemoteFXSynth3DVSCVMDevice
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataMicrosoftWindowsRemoteDesktopServicesRemoteFXSynth3dvscRemoteFXSynth3DVSCVMDevice
        Case wmiEnumClassWin32PerfFormattedDataMicrosoftWindowsRemoteDesktopServicesRemoteFXSynth3dvscRemoteFXSynth3DVSCVMTransportChannel
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataMicrosoftWindowsRemoteDesktopServicesRemoteFXSynth3dvscRemoteFXSynth3DVSCVMTransportChannel
        Case wmiEnumClassWin32PerfFormattedDataMSDTCDistributedTransactionCoordinator
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataMSDTCDistributedTransactionCoordinator
        Case wmiEnumClassWin32PerfFormattedDataMSDTCBridge3000MSDTCBridge3000
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataMSDTCBridge3000MSDTCBridge3000
        Case wmiEnumClassWin32PerfFormattedDataMSDTCBridge4000MSDTCBridge4000
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataMSDTCBridge4000MSDTCBridge4000
        Case wmiEnumClassWin32PerfFormattedDataNETCLRDataNETCLRData
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataNETCLRDataNETCLRData
        Case wmiEnumClassWin32PerfFormattedDataNETCLRNetworkingNETCLRNetworking
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataNETCLRNetworkingNETCLRNetworking
        Case wmiEnumClassWin32PerfFormattedDataNETCLRNetworking4000NETCLRNetworking4000
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataNETCLRNetworking4000NETCLRNetworking4000
        Case wmiEnumClassWin32PerfFormattedDataNETDataProviderforOracleNETDataProviderforOracle
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataNETDataProviderforOracleNETDataProviderforOracle
        Case wmiEnumClassWin32PerfFormattedDataNETDataProviderforSqlServerNETDataProviderforSqlServer
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataNETDataProviderforSqlServerNETDataProviderforSqlServer
        Case wmiEnumClassWin32PerfFormattedDataNETFrameworkNETCLRExceptions
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataNETFrameworkNETCLRExceptions
        Case wmiEnumClassWin32PerfFormattedDataNETFrameworkNETCLRInterop
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataNETFrameworkNETCLRInterop
        Case wmiEnumClassWin32PerfFormattedDataNETFrameworkNETCLRJit
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataNETFrameworkNETCLRJit
        Case wmiEnumClassWin32PerfFormattedDataNETFrameworkNETCLRLoading
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataNETFrameworkNETCLRLoading
        Case wmiEnumClassWin32PerfFormattedDataNETFrameworkNETCLRLocksAndThreads
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataNETFrameworkNETCLRLocksAndThreads
        Case wmiEnumClassWin32PerfFormattedDataNETFrameworkNETCLRMemory
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataNETFrameworkNETCLRMemory
        Case wmiEnumClassWin32PerfFormattedDataNETFrameworkNETCLRRemoting
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataNETFrameworkNETCLRRemoting
        Case wmiEnumClassWin32PerfFormattedDataNETFrameworkNETCLRSecurity
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataNETFrameworkNETCLRSecurity
        Case wmiEnumClassWin32PerfFormattedDataNETMemoryCache40NETMemoryCache40
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataNETMemoryCache40NETMemoryCache40
        Case wmiEnumClassWin32PerfFormattedDataNvspNicStatsHyperVVirtualNetworkAdapter
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataNvspNicStatsHyperVVirtualNetworkAdapter
        Case wmiEnumClassWin32PerfFormattedDataNvspPortStatsHyperVVirtualSwitchPort
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataNvspPortStatsHyperVVirtualSwitchPort
        Case wmiEnumClassWin32PerfFormattedDataNvspSwitchStatsHyperVVirtualSwitch
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataNvspSwitchStatsHyperVVirtualSwitch
        Case wmiEnumClassWin32PerfFormattedDataOfflineFilesClientSideCaching
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataOfflineFilesClientSideCaching
        Case wmiEnumClassWin32PerfFormattedDataOfflineFilesOfflineFiles
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataOfflineFilesOfflineFiles
        Case wmiEnumClassWin32PerfFormattedDataPeerDistSvcBranchCache
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataPeerDistSvcBranchCache
        Case wmiEnumClassWin32PerfFormattedDataPeerNameResolutionProtocolPerfPeerNameResolutionProtocol
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataPeerNameResolutionProtocolPerfPeerNameResolutionProtocol
        Case wmiEnumClassWin32PerfFormattedDataPerfDiskLogicalDisk
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataPerfDiskLogicalDisk
        Case wmiEnumClassWin32PerfFormattedDataPerfDiskPhysicalDisk
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataPerfDiskPhysicalDisk
        Case wmiEnumClassWin32PerfFormattedDataPerfNetBrowser
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataPerfNetBrowser
        Case wmiEnumClassWin32PerfFormattedDataPerfNetRedirector
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataPerfNetRedirector
        Case wmiEnumClassWin32PerfFormattedDataPerfNetServer
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataPerfNetServer
        Case wmiEnumClassWin32PerfFormattedDataPerfNetServerWorkQueues
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataPerfNetServerWorkQueues
        Case wmiEnumClassWin32PerfFormattedDataPerfOSCache
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataPerfOSCache
        Case wmiEnumClassWin32PerfFormattedDataPerfOSMemory
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataPerfOSMemory
        Case wmiEnumClassWin32PerfFormattedDataPerfOSNUMANodeMemory
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataPerfOSNUMANodeMemory
        Case wmiEnumClassWin32PerfFormattedDataPerfOSObjects
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataPerfOSObjects
        Case wmiEnumClassWin32PerfFormattedDataPerfOSPagingFile
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataPerfOSPagingFile
        Case wmiEnumClassWin32PerfFormattedDataPerfOSProcessor
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataPerfOSProcessor
        Case wmiEnumClassWin32PerfFormattedDataPerfOSSystem
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataPerfOSSystem
        Case wmiEnumClassWin32PerfFormattedDataPerfProcFullImageCostly
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataPerfProcFullImageCostly
        Case wmiEnumClassWin32PerfFormattedDataPerfProcImageCostly
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataPerfProcImageCostly
        Case wmiEnumClassWin32PerfFormattedDataPerfProcJobObject
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataPerfProcJobObject
        Case wmiEnumClassWin32PerfFormattedDataPerfProcJobObjectDetails
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataPerfProcJobObjectDetails
        Case wmiEnumClassWin32PerfFormattedDataPerfProcProcess
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataPerfProcProcess
        Case wmiEnumClassWin32PerfFormattedDataPerfProcProcessAddressSpaceCostly
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataPerfProcProcessAddressSpaceCostly
        Case wmiEnumClassWin32PerfFormattedDataPerfProcThread
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataPerfProcThread
        Case wmiEnumClassWin32PerfFormattedDataPerfProcThreadDetailsCostly
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataPerfProcThreadDetailsCostly
        Case wmiEnumClassWin32PerfFormattedDataPowerMeterCounterEnergyMeter
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataPowerMeterCounterEnergyMeter
        Case wmiEnumClassWin32PerfFormattedDataPowerMeterCounterPowerMeter
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataPowerMeterCounterPowerMeter
        Case wmiEnumClassWin32PerfFormattedDatardyboostReadyBoostCache
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDatardyboostReadyBoostCache
        Case wmiEnumClassWin32PerfFormattedDataRemoteAccessRASPort
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataRemoteAccessRASPort
        Case wmiEnumClassWin32PerfFormattedDataRemoteAccessRASTotal
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataRemoteAccessRASTotal
        Case wmiEnumClassWin32PerfFormattedDataRemotePerfProviderHyperVVMRemoting
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataRemotePerfProviderHyperVVMRemoting
        Case wmiEnumClassWin32PerfFormattedDataServiceModel4000ServiceModelEndpoint4000
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataServiceModel4000ServiceModelEndpoint4000
        Case wmiEnumClassWin32PerfFormattedDataServiceModel4000ServiceModelOperation4000
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataServiceModel4000ServiceModelOperation4000
        Case wmiEnumClassWin32PerfFormattedDataServiceModel4000ServiceModelService4000
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataServiceModel4000ServiceModelService4000
        Case wmiEnumClassWin32PerfFormattedDataServiceModelEndpoint3000ServiceModelEndpoint3000
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataServiceModelEndpoint3000ServiceModelEndpoint3000
        Case wmiEnumClassWin32PerfFormattedDataServiceModelOperation3000ServiceModelOperation3000
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataServiceModelOperation3000ServiceModelOperation3000
        Case wmiEnumClassWin32PerfFormattedDataServiceModelService3000ServiceModelService3000
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataServiceModelService3000ServiceModelService3000
        Case wmiEnumClassWin32PerfFormattedDataSMSvcHost3000SMSvcHost3000
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataSMSvcHost3000SMSvcHost3000
        Case wmiEnumClassWin32PerfFormattedDataSMSvcHost4000SMSvcHost4000
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataSMSvcHost4000SMSvcHost4000
        Case wmiEnumClassWin32PerfFormattedDataSpoolerPrintQueue
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataSpoolerPrintQueue
        Case wmiEnumClassWin32PerfFormattedDataStorageStatsHyperVVirtualStorageDevice
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataStorageStatsHyperVVirtualStorageDevice
        Case wmiEnumClassWin32PerfFormattedDataTapiSrvTelephony
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataTapiSrvTelephony
        Case wmiEnumClassWin32PerfFormattedDataTBSTBScounters
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataTBSTBScounters
        Case wmiEnumClassWin32PerfFormattedDataTcpipICMP
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataTcpipICMP
        Case wmiEnumClassWin32PerfFormattedDataTcpipICMPv6
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataTcpipICMPv6
        Case wmiEnumClassWin32PerfFormattedDataTcpipIPv4
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataTcpipIPv4
        Case wmiEnumClassWin32PerfFormattedDataTcpipIPv6
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataTcpipIPv6
        Case wmiEnumClassWin32PerfFormattedDataTcpipNBTConnection
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataTcpipNBTConnection
        Case wmiEnumClassWin32PerfFormattedDataTcpipNetworkAdapter
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataTcpipNetworkAdapter
        Case wmiEnumClassWin32PerfFormattedDataTcpipNetworkInterface
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataTcpipNetworkInterface
        Case wmiEnumClassWin32PerfFormattedDataTcpipTCPv4
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataTcpipTCPv4
        Case wmiEnumClassWin32PerfFormattedDataTcpipTCPv6
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataTcpipTCPv6
        Case wmiEnumClassWin32PerfFormattedDataTcpipUDPv4
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataTcpipUDPv4
        Case wmiEnumClassWin32PerfFormattedDataTcpipUDPv6
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataTcpipUDPv6
        Case wmiEnumClassWin32PerfFormattedDataTCPIPCountersTCPIPPerformanceDiagnostics
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataTCPIPCountersTCPIPPerformanceDiagnostics
        Case wmiEnumClassWin32PerfFormattedDataTermServiceTerminalServicesSession
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataTermServiceTerminalServicesSession
        Case wmiEnumClassWin32PerfFormattedDataUGathererSearchGathererProjects
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataUGathererSearchGathererProjects
        Case wmiEnumClassWin32PerfFormattedDataUGTHRSVCSearchGatherer
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataUGTHRSVCSearchGatherer
        Case wmiEnumClassWin32PerfFormattedDatausbhubUSB
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDatausbhubUSB
        Case wmiEnumClassWin32PerfFormattedDataVidPerfProviderHyperVVMVidNumaNode
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataVidPerfProviderHyperVVMVidNumaNode
        Case wmiEnumClassWin32PerfFormattedDataVidPerfProviderHyperVVMVidPartition
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataVidPerfProviderHyperVVMVidPartition
        Case wmiEnumClassWin32PerfFormattedDataVmbusStatsHyperVVirtualMachineBus
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataVmbusStatsHyperVVirtualMachineBus
        Case wmiEnumClassWin32PerfFormattedDataVmmsVirtualMachineStatsHyperVVirtualMachineHealthSummary
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataVmmsVirtualMachineStatsHyperVVirtualMachineHealthSummary
        Case wmiEnumClassWin32PerfFormattedDataVmmsVirtualMachineStatsHyperVVirtualMachineSummary
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataVmmsVirtualMachineStatsHyperVVirtualMachineSummary
        Case wmiEnumClassWin32PerfFormattedDataVmTaskManagerStatsHyperVTaskManagerDetail
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataVmTaskManagerStatsHyperVTaskManagerDetail
        Case wmiEnumClassWin32PerfFormattedDataW3SVCWebService
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataW3SVCWebService
        Case wmiEnumClassWin32PerfFormattedDataW3SVCWebServiceCache
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataW3SVCWebServiceCache
        Case wmiEnumClassWin32PerfFormattedDataW3SVCW3WPCounterProviderW3SVCW3WP
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataW3SVCW3WPCounterProviderW3SVCW3WP
        Case wmiEnumClassWin32PerfFormattedDataWASW3WPCounterProviderWASW3WP
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataWASW3WPCounterProviderWASW3WP
        Case wmiEnumClassWin32PerfFormattedDataWindowsMediaPlayerWindowsMediaPlayerMetadata
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataWindowsMediaPlayerWindowsMediaPlayerMetadata
        Case wmiEnumClassWin32PerfFormattedDataWindowsWorkflowFoundation3000WindowsWorkflowFoundation
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataWindowsWorkflowFoundation3000WindowsWorkflowFoundation
        Case wmiEnumClassWin32PerfFormattedDataWindowsWorkflowFoundation4000WFSystemWorkflow4000
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataWindowsWorkflowFoundation4000WFSystemWorkflow4000
        Case wmiEnumClassWin32PerfFormattedDataWorkflowServiceHost4000WorkflowServiceHost4000
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataWorkflowServiceHost4000WorkflowServiceHost4000
        Case wmiEnumClassWin32PerfFormattedDataWSearchIdxPiSearchIndexer
            SelectWmiClassName = WmiClassNameWin32PerfFormattedDataWSearchIdxPiSearchIndexer
        Case wmiEnumClassWin32PerfRawData
            SelectWmiClassName = WmiClassNameWin32PerfRawData
        Case wmiEnumClassWin32PerfRawDataAFDCountersMicrosoftWinsockBSP
            SelectWmiClassName = WmiClassNameWin32PerfRawDataAFDCountersMicrosoftWinsockBSP
        Case wmiEnumClassWin32PerfRawDataAPPPOOLCountersProviderAPPPOOLWAS
            SelectWmiClassName = WmiClassNameWin32PerfRawDataAPPPOOLCountersProviderAPPPOOLWAS
        Case wmiEnumClassWin32PerfRawDataASPActiveServerPages
            SelectWmiClassName = WmiClassNameWin32PerfRawDataASPActiveServerPages
        Case wmiEnumClassWin32PerfRawDataASPNETASPNET
            SelectWmiClassName = WmiClassNameWin32PerfRawDataASPNETASPNET
        Case wmiEnumClassWin32PerfRawDataASPNETASPNETApplications
            SelectWmiClassName = WmiClassNameWin32PerfRawDataASPNETASPNETApplications
        Case wmiEnumClassWin32PerfRawDataASPNET2050727ASPNETAppsv2050727
            SelectWmiClassName = WmiClassNameWin32PerfRawDataASPNET2050727ASPNETAppsv2050727
        Case wmiEnumClassWin32PerfRawDataASPNET2050727ASPNETv2050727
            SelectWmiClassName = WmiClassNameWin32PerfRawDataASPNET2050727ASPNETv2050727
        Case wmiEnumClassWin32PerfRawDataASPNET4030319ASPNETAppsv4030319
            SelectWmiClassName = WmiClassNameWin32PerfRawDataASPNET4030319ASPNETAppsv4030319
        Case wmiEnumClassWin32PerfRawDataASPNET4030319ASPNETv4030319
            SelectWmiClassName = WmiClassNameWin32PerfRawDataASPNET4030319ASPNETv4030319
        Case wmiEnumClassWin32PerfRawDataaspnetstateASPNETStateService
            SelectWmiClassName = WmiClassNameWin32PerfRawDataaspnetstateASPNETStateService
        Case wmiEnumClassWin32PerfRawDataAuthorizationManagerAuthorizationManagerApplications
            SelectWmiClassName = WmiClassNameWin32PerfRawDataAuthorizationManagerAuthorizationManagerApplications
        Case wmiEnumClassWin32PerfRawDataBalancerStatsHyperVDynamicMemoryBalancer
            SelectWmiClassName = WmiClassNameWin32PerfRawDataBalancerStatsHyperVDynamicMemoryBalancer
        Case wmiEnumClassWin32PerfRawDataBalancerStatsHyperVDynamicMemoryVM
            SelectWmiClassName = WmiClassNameWin32PerfRawDataBalancerStatsHyperVDynamicMemoryVM
        Case wmiEnumClassWin32PerfRawDataBITSBITSNetUtilization
            SelectWmiClassName = WmiClassNameWin32PerfRawDataBITSBITSNetUtilization
        Case wmiEnumClassWin32PerfRawDataCountersDNS64Global
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersDNS64Global
        Case wmiEnumClassWin32PerfRawDataCountersEventTracingforWindows
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersEventTracingforWindows
        Case wmiEnumClassWin32PerfRawDataCountersEventTracingforWindowsSession
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersEventTracingforWindowsSession
        Case wmiEnumClassWin32PerfRawDataCountersFileSystemDiskActivity
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersFileSystemDiskActivity
        Case wmiEnumClassWin32PerfRawDataCountersGenericIKEv1AuthIPandIKEv2
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersGenericIKEv1AuthIPandIKEv2
        Case wmiEnumClassWin32PerfRawDataCountersHTTPService
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersHTTPService
        Case wmiEnumClassWin32PerfRawDataCountersHTTPServiceRequestQueues
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersHTTPServiceRequestQueues
        Case wmiEnumClassWin32PerfRawDataCountersHTTPServiceUrlGroups
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersHTTPServiceUrlGroups
        Case wmiEnumClassWin32PerfRawDataCountersHyperVDynamicMemoryIntegrationService
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersHyperVDynamicMemoryIntegrationService
        Case wmiEnumClassWin32PerfRawDataCountersHyperVVirtualMachineBusPipes
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersHyperVVirtualMachineBusPipes
        Case wmiEnumClassWin32PerfRawDataCountersIPHTTPSGlobal
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersIPHTTPSGlobal
        Case wmiEnumClassWin32PerfRawDataCountersIPHTTPSSession
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersIPHTTPSSession
        Case wmiEnumClassWin32PerfRawDataCountersIPsecAuthIPIPv4
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersIPsecAuthIPIPv4
        Case wmiEnumClassWin32PerfRawDataCountersIPsecAuthIPIPv6
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersIPsecAuthIPIPv6
        Case wmiEnumClassWin32PerfRawDataCountersIPsecConnections
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersIPsecConnections
        Case wmiEnumClassWin32PerfRawDataCountersIPsecDoSProtection
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersIPsecDoSProtection
        Case wmiEnumClassWin32PerfRawDataCountersIPsecDriver
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersIPsecDriver
        Case wmiEnumClassWin32PerfRawDataCountersIPsecIKEv1IPv4
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersIPsecIKEv1IPv4
        Case wmiEnumClassWin32PerfRawDataCountersIPsecIKEv1IPv6
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersIPsecIKEv1IPv6
        Case wmiEnumClassWin32PerfRawDataCountersIPsecIKEv2IPv4
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersIPsecIKEv2IPv4
        Case wmiEnumClassWin32PerfRawDataCountersIPsecIKEv2IPv6
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersIPsecIKEv2IPv6
        Case wmiEnumClassWin32PerfRawDataCountersNetlogon
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersNetlogon
        Case wmiEnumClassWin32PerfRawDataCountersNetworkQoSPolicy
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersNetworkQoSPolicy
        Case wmiEnumClassWin32PerfRawDataCountersPacerFlow
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersPacerFlow
        Case wmiEnumClassWin32PerfRawDataCountersPacerPipe
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersPacerPipe
        Case wmiEnumClassWin32PerfRawDataCountersPacketDirectECUtilization
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersPacketDirectECUtilization
        Case wmiEnumClassWin32PerfRawDataCountersPacketDirectQueueDepth
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersPacketDirectQueueDepth
        Case wmiEnumClassWin32PerfRawDataCountersPacketDirectReceiveCounters
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersPacketDirectReceiveCounters
        Case wmiEnumClassWin32PerfRawDataCountersPacketDirectReceiveFilters
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersPacketDirectReceiveFilters
        Case wmiEnumClassWin32PerfRawDataCountersPacketDirectTransmitCounters
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersPacketDirectTransmitCounters
        Case wmiEnumClassWin32PerfRawDataCountersPerProcessorNetworkActivityCycles
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersPerProcessorNetworkActivityCycles
        Case wmiEnumClassWin32PerfRawDataCountersPerProcessorNetworkInterfaceCardActivity
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersPerProcessorNetworkInterfaceCardActivity
        Case wmiEnumClassWin32PerfRawDataCountersPhysicalNetworkInterfaceCardActivity
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersPhysicalNetworkInterfaceCardActivity
        Case wmiEnumClassWin32PerfRawDataCountersPowerShellWorkflow
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersPowerShellWorkflow
        Case wmiEnumClassWin32PerfRawDataCountersProcessorInformation
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersProcessorInformation
        Case wmiEnumClassWin32PerfRawDataCountersRDMAActivity
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersRDMAActivity
        Case wmiEnumClassWin32PerfRawDataCountersRemoteFXGraphics
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersRemoteFXGraphics
        Case wmiEnumClassWin32PerfRawDataCountersRemoteFXNetwork
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersRemoteFXNetwork
        Case wmiEnumClassWin32PerfRawDataCountersSMBClientShares
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersSMBClientShares
        Case wmiEnumClassWin32PerfRawDataCountersSMBServer
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersSMBServer
        Case wmiEnumClassWin32PerfRawDataCountersSMBServerSessions
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersSMBServerSessions
        Case wmiEnumClassWin32PerfRawDataCountersSMBServerShares
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersSMBServerShares
        Case wmiEnumClassWin32PerfRawDataCountersStorageSpacesTier
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersStorageSpacesTier
        Case wmiEnumClassWin32PerfRawDataCountersStorageSpacesWriteCache
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersStorageSpacesWriteCache
        Case wmiEnumClassWin32PerfRawDataCountersSynchronization
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersSynchronization
        Case wmiEnumClassWin32PerfRawDataCountersSynchronizationNuma
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersSynchronizationNuma
        Case wmiEnumClassWin32PerfRawDataCountersTeredoClient
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersTeredoClient
        Case wmiEnumClassWin32PerfRawDataCountersTeredoRelay
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersTeredoRelay
        Case wmiEnumClassWin32PerfRawDataCountersTeredoServer
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersTeredoServer
        Case wmiEnumClassWin32PerfRawDataCountersThermalZoneInformation
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersThermalZoneInformation
        Case wmiEnumClassWin32PerfRawDataCountersWFP
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersWFP
        Case wmiEnumClassWin32PerfRawDataCountersWFPv4
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersWFPv4
        Case wmiEnumClassWin32PerfRawDataCountersWFPv6
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersWFPv6
        Case wmiEnumClassWin32PerfRawDataCountersWSManQuotaStatistics
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersWSManQuotaStatistics
        Case wmiEnumClassWin32PerfRawDataCountersXHCICommonBuffer
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersXHCICommonBuffer
        Case wmiEnumClassWin32PerfRawDataCountersXHCIInterrupter
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersXHCIInterrupter
        Case wmiEnumClassWin32PerfRawDataCountersXHCITransferRing
            SelectWmiClassName = WmiClassNameWin32PerfRawDataCountersXHCITransferRing
        Case wmiEnumClassWin32PerfRawDataDdmCounterProviderRAS
            SelectWmiClassName = WmiClassNameWin32PerfRawDataDdmCounterProviderRAS
        Case wmiEnumClassWin32PerfRawDataDeliveryOptimizationDeliveryOptimizationSwarm
            SelectWmiClassName = WmiClassNameWin32PerfRawDataDeliveryOptimizationDeliveryOptimizationSwarm
        Case wmiEnumClassWin32PerfRawDataDistributedRoutingTablePerfDistributedRoutingTable
            SelectWmiClassName = WmiClassNameWin32PerfRawDataDistributedRoutingTablePerfDistributedRoutingTable
        Case wmiEnumClassWin32PerfRawDataESENTDatabase
            SelectWmiClassName = WmiClassNameWin32PerfRawDataESENTDatabase
        Case wmiEnumClassWin32PerfRawDataESENTDatabaseInstances
            SelectWmiClassName = WmiClassNameWin32PerfRawDataESENTDatabaseInstances
        Case wmiEnumClassWin32PerfRawDataESENTDatabaseTableClasses
            SelectWmiClassName = WmiClassNameWin32PerfRawDataESENTDatabaseTableClasses
        Case wmiEnumClassWin32PerfRawDataEthernetPerfProviderHyperVLegacyNetworkAdapter
            SelectWmiClassName = WmiClassNameWin32PerfRawDataEthernetPerfProviderHyperVLegacyNetworkAdapter
        Case wmiEnumClassWin32PerfRawDataFaxServiceFaxService
            SelectWmiClassName = WmiClassNameWin32PerfRawDataFaxServiceFaxService
        Case wmiEnumClassWin32PerfRawDataftpsvcMicrosoftFTPService
            SelectWmiClassName = WmiClassNameWin32PerfRawDataftpsvcMicrosoftFTPService
        Case wmiEnumClassWin32PerfRawDataGmoPerfProviderHyperVVMSaveSnapshotandRestore
            SelectWmiClassName = WmiClassNameWin32PerfRawDataGmoPerfProviderHyperVVMSaveSnapshotandRestore
        Case wmiEnumClassWin32PerfRawDataHvStatsHyperVHypervisor
            SelectWmiClassName = WmiClassNameWin32PerfRawDataHvStatsHyperVHypervisor
        Case wmiEnumClassWin32PerfRawDataHvStatsHyperVHypervisorLogicalProcessor
            SelectWmiClassName = WmiClassNameWin32PerfRawDataHvStatsHyperVHypervisorLogicalProcessor
        Case wmiEnumClassWin32PerfRawDataHvStatsHyperVHypervisorPartition
            SelectWmiClassName = WmiClassNameWin32PerfRawDataHvStatsHyperVHypervisorPartition
        Case wmiEnumClassWin32PerfRawDataHvStatsHyperVHypervisorRootPartition
            SelectWmiClassName = WmiClassNameWin32PerfRawDataHvStatsHyperVHypervisorRootPartition
        Case wmiEnumClassWin32PerfRawDataHvStatsHyperVHypervisorRootVirtualProcessor
            SelectWmiClassName = WmiClassNameWin32PerfRawDataHvStatsHyperVHypervisorRootVirtualProcessor
        Case wmiEnumClassWin32PerfRawDataHvStatsHyperVHypervisorVirtualProcessor
            SelectWmiClassName = WmiClassNameWin32PerfRawDataHvStatsHyperVHypervisorVirtualProcessor
        Case wmiEnumClassWin32PerfRawDataIdePerfProviderHyperVVirtualIDEController
            SelectWmiClassName = WmiClassNameWin32PerfRawDataIdePerfProviderHyperVVirtualIDEController
        Case wmiEnumClassWin32PerfRawDataLocalSessionManagerTerminalServices
            SelectWmiClassName = WmiClassNameWin32PerfRawDataLocalSessionManagerTerminalServices
        Case wmiEnumClassWin32PerfRawDataLsaSecurityPerProcessStatistics
            SelectWmiClassName = WmiClassNameWin32PerfRawDataLsaSecurityPerProcessStatistics
        Case wmiEnumClassWin32PerfRawDataLsaSecuritySystemWideStatistics
            SelectWmiClassName = WmiClassNameWin32PerfRawDataLsaSecuritySystemWideStatistics
        Case wmiEnumClassWin32PerfRawDataMicrosoftWindowsBitLockerDriverCountersProviderBitLocker
            SelectWmiClassName = WmiClassNameWin32PerfRawDataMicrosoftWindowsBitLockerDriverCountersProviderBitLocker
        Case wmiEnumClassWin32PerfRawDataMicrosoftWindowsRemoteDesktopServicesRemoteFXSynth3dvscRemoteFXSynth3DVSCVMDevice
            SelectWmiClassName = WmiClassNameWin32PerfRawDataMicrosoftWindowsRemoteDesktopServicesRemoteFXSynth3dvscRemoteFXSynth3DVSCVMDevice
        Case wmiEnumClassWin32PerfRawDataMicrosoftWindowsRemoteDesktopServicesRemoteFXSynth3dvscRemoteFXSynth3DVSCVMTransportChannel
            SelectWmiClassName = WmiClassNameWin32PerfRawDataMicrosoftWindowsRemoteDesktopServicesRemoteFXSynth3dvscRemoteFXSynth3DVSCVMTransportChannel
        Case wmiEnumClassWin32PerfRawDataMSDTCDistributedTransactionCoordinator
            SelectWmiClassName = WmiClassNameWin32PerfRawDataMSDTCDistributedTransactionCoordinator
        Case wmiEnumClassWin32PerfRawDataMSDTCBridge3000MSDTCBridge3000
            SelectWmiClassName = WmiClassNameWin32PerfRawDataMSDTCBridge3000MSDTCBridge3000
        Case wmiEnumClassWin32PerfRawDataMSDTCBridge4000MSDTCBridge4000
            SelectWmiClassName = WmiClassNameWin32PerfRawDataMSDTCBridge4000MSDTCBridge4000
        Case wmiEnumClassWin32PerfRawDataNETCLRDataNETCLRData
            SelectWmiClassName = WmiClassNameWin32PerfRawDataNETCLRDataNETCLRData
        Case wmiEnumClassWin32PerfRawDataNETCLRNetworkingNETCLRNetworking
            SelectWmiClassName = WmiClassNameWin32PerfRawDataNETCLRNetworkingNETCLRNetworking
        Case wmiEnumClassWin32PerfRawDataNETCLRNetworking4000NETCLRNetworking4000
            SelectWmiClassName = WmiClassNameWin32PerfRawDataNETCLRNetworking4000NETCLRNetworking4000
        Case wmiEnumClassWin32PerfRawDataNETDataProviderforOracleNETDataProviderforOracle
            SelectWmiClassName = WmiClassNameWin32PerfRawDataNETDataProviderforOracleNETDataProviderforOracle
        Case wmiEnumClassWin32PerfRawDataNETDataProviderforSqlServerNETDataProviderforSqlServer
            SelectWmiClassName = WmiClassNameWin32PerfRawDataNETDataProviderforSqlServerNETDataProviderforSqlServer
        Case wmiEnumClassWin32PerfRawDataNETFrameworkNETCLRExceptions
            SelectWmiClassName = WmiClassNameWin32PerfRawDataNETFrameworkNETCLRExceptions
        Case wmiEnumClassWin32PerfRawDataNETFrameworkNETCLRInterop
            SelectWmiClassName = WmiClassNameWin32PerfRawDataNETFrameworkNETCLRInterop
        Case wmiEnumClassWin32PerfRawDataNETFrameworkNETCLRJit
            SelectWmiClassName = WmiClassNameWin32PerfRawDataNETFrameworkNETCLRJit
        Case wmiEnumClassWin32PerfRawDataNETFrameworkNETCLRLoading
            SelectWmiClassName = WmiClassNameWin32PerfRawDataNETFrameworkNETCLRLoading
        Case wmiEnumClassWin32PerfRawDataNETFrameworkNETCLRLocksAndThreads
            SelectWmiClassName = WmiClassNameWin32PerfRawDataNETFrameworkNETCLRLocksAndThreads
        Case wmiEnumClassWin32PerfRawDataNETFrameworkNETCLRMemory
            SelectWmiClassName = WmiClassNameWin32PerfRawDataNETFrameworkNETCLRMemory
        Case wmiEnumClassWin32PerfRawDataNETFrameworkNETCLRRemoting
            SelectWmiClassName = WmiClassNameWin32PerfRawDataNETFrameworkNETCLRRemoting
        Case wmiEnumClassWin32PerfRawDataNETFrameworkNETCLRSecurity
            SelectWmiClassName = WmiClassNameWin32PerfRawDataNETFrameworkNETCLRSecurity
        Case wmiEnumClassWin32PerfRawDataNETMemoryCache40NETMemoryCache40
            SelectWmiClassName = WmiClassNameWin32PerfRawDataNETMemoryCache40NETMemoryCache40
        Case wmiEnumClassWin32PerfRawDataNvspNicStatsHyperVVirtualNetworkAdapter
            SelectWmiClassName = WmiClassNameWin32PerfRawDataNvspNicStatsHyperVVirtualNetworkAdapter
        Case wmiEnumClassWin32PerfRawDataNvspPortStatsHyperVVirtualSwitchPort
            SelectWmiClassName = WmiClassNameWin32PerfRawDataNvspPortStatsHyperVVirtualSwitchPort
        Case wmiEnumClassWin32PerfRawDataNvspSwitchStatsHyperVVirtualSwitch
            SelectWmiClassName = WmiClassNameWin32PerfRawDataNvspSwitchStatsHyperVVirtualSwitch
        Case wmiEnumClassWin32PerfRawDataOfflineFilesClientSideCaching
            SelectWmiClassName = WmiClassNameWin32PerfRawDataOfflineFilesClientSideCaching
        Case wmiEnumClassWin32PerfRawDataOfflineFilesOfflineFiles
            SelectWmiClassName = WmiClassNameWin32PerfRawDataOfflineFilesOfflineFiles
        Case wmiEnumClassWin32PerfRawDataPeerDistSvcBranchCache
            SelectWmiClassName = WmiClassNameWin32PerfRawDataPeerDistSvcBranchCache
        Case wmiEnumClassWin32PerfRawDataPeerNameResolutionProtocolPerfPeerNameResolutionProtocol
            SelectWmiClassName = WmiClassNameWin32PerfRawDataPeerNameResolutionProtocolPerfPeerNameResolutionProtocol
        Case wmiEnumClassWin32PerfRawDataPerfDiskLogicalDisk
            SelectWmiClassName = WmiClassNameWin32PerfRawDataPerfDiskLogicalDisk
        Case wmiEnumClassWin32PerfRawDataPerfDiskPhysicalDisk
            SelectWmiClassName = WmiClassNameWin32PerfRawDataPerfDiskPhysicalDisk
        Case wmiEnumClassWin32PerfRawDataPerfNetBrowser
            SelectWmiClassName = WmiClassNameWin32PerfRawDataPerfNetBrowser
        Case wmiEnumClassWin32PerfRawDataPerfNetRedirector
            SelectWmiClassName = WmiClassNameWin32PerfRawDataPerfNetRedirector
        Case wmiEnumClassWin32PerfRawDataPerfNetServer
            SelectWmiClassName = WmiClassNameWin32PerfRawDataPerfNetServer
        Case wmiEnumClassWin32PerfRawDataPerfNetServerWorkQueues
            SelectWmiClassName = WmiClassNameWin32PerfRawDataPerfNetServerWorkQueues
        Case wmiEnumClassWin32PerfRawDataPerfOSCache
            SelectWmiClassName = WmiClassNameWin32PerfRawDataPerfOSCache
        Case wmiEnumClassWin32PerfRawDataPerfOSMemory
            SelectWmiClassName = WmiClassNameWin32PerfRawDataPerfOSMemory
        Case wmiEnumClassWin32PerfRawDataPerfOSNUMANodeMemory
            SelectWmiClassName = WmiClassNameWin32PerfRawDataPerfOSNUMANodeMemory
        Case wmiEnumClassWin32PerfRawDataPerfOSObjects
            SelectWmiClassName = WmiClassNameWin32PerfRawDataPerfOSObjects
        Case wmiEnumClassWin32PerfRawDataPerfOSPagingFile
            SelectWmiClassName = WmiClassNameWin32PerfRawDataPerfOSPagingFile
        Case wmiEnumClassWin32PerfRawDataPerfOSProcessor
            SelectWmiClassName = WmiClassNameWin32PerfRawDataPerfOSProcessor
        Case wmiEnumClassWin32PerfRawDataPerfOSSystem
            SelectWmiClassName = WmiClassNameWin32PerfRawDataPerfOSSystem
        Case wmiEnumClassWin32PerfRawDataPerfProcFullImageCostly
            SelectWmiClassName = WmiClassNameWin32PerfRawDataPerfProcFullImageCostly
        Case wmiEnumClassWin32PerfRawDataPerfProcImageCostly
            SelectWmiClassName = WmiClassNameWin32PerfRawDataPerfProcImageCostly
        Case wmiEnumClassWin32PerfRawDataPerfProcJobObject
            SelectWmiClassName = WmiClassNameWin32PerfRawDataPerfProcJobObject
        Case wmiEnumClassWin32PerfRawDataPerfProcJobObjectDetails
            SelectWmiClassName = WmiClassNameWin32PerfRawDataPerfProcJobObjectDetails
        Case wmiEnumClassWin32PerfRawDataPerfProcProcess
            SelectWmiClassName = WmiClassNameWin32PerfRawDataPerfProcProcess
        Case wmiEnumClassWin32PerfRawDataPerfProcProcessAddressSpaceCostly
            SelectWmiClassName = WmiClassNameWin32PerfRawDataPerfProcProcessAddressSpaceCostly
        Case wmiEnumClassWin32PerfRawDataPerfProcThread
            SelectWmiClassName = WmiClassNameWin32PerfRawDataPerfProcThread
        Case wmiEnumClassWin32PerfRawDataPerfProcThreadDetailsCostly
            SelectWmiClassName = WmiClassNameWin32PerfRawDataPerfProcThreadDetailsCostly
        Case wmiEnumClassWin32PerfRawDataPowerMeterCounterEnergyMeter
            SelectWmiClassName = WmiClassNameWin32PerfRawDataPowerMeterCounterEnergyMeter
        Case wmiEnumClassWin32PerfRawDataPowerMeterCounterPowerMeter
            SelectWmiClassName = WmiClassNameWin32PerfRawDataPowerMeterCounterPowerMeter
        Case wmiEnumClassWin32PerfRawDatardyboostReadyBoostCache
            SelectWmiClassName = WmiClassNameWin32PerfRawDatardyboostReadyBoostCache
        Case wmiEnumClassWin32PerfRawDataRemoteAccessRASPort
            SelectWmiClassName = WmiClassNameWin32PerfRawDataRemoteAccessRASPort
        Case wmiEnumClassWin32PerfRawDataRemoteAccessRASTotal
            SelectWmiClassName = WmiClassNameWin32PerfRawDataRemoteAccessRASTotal
        Case wmiEnumClassWin32PerfRawDataRemotePerfProviderHyperVVMRemoting
            SelectWmiClassName = WmiClassNameWin32PerfRawDataRemotePerfProviderHyperVVMRemoting
        Case wmiEnumClassWin32PerfRawDataServiceModel4000ServiceModelEndpoint4000
            SelectWmiClassName = WmiClassNameWin32PerfRawDataServiceModel4000ServiceModelEndpoint4000
        Case wmiEnumClassWin32PerfRawDataServiceModel4000ServiceModelOperation4000
            SelectWmiClassName = WmiClassNameWin32PerfRawDataServiceModel4000ServiceModelOperation4000
        Case wmiEnumClassWin32PerfRawDataServiceModel4000ServiceModelService4000
            SelectWmiClassName = WmiClassNameWin32PerfRawDataServiceModel4000ServiceModelService4000
        Case wmiEnumClassWin32PerfRawDataServiceModelEndpoint3000ServiceModelEndpoint3000
            SelectWmiClassName = WmiClassNameWin32PerfRawDataServiceModelEndpoint3000ServiceModelEndpoint3000
        Case wmiEnumClassWin32PerfRawDataServiceModelOperation3000ServiceModelOperation3000
            SelectWmiClassName = WmiClassNameWin32PerfRawDataServiceModelOperation3000ServiceModelOperation3000
        Case wmiEnumClassWin32PerfRawDataServiceModelService3000ServiceModelService3000
            SelectWmiClassName = WmiClassNameWin32PerfRawDataServiceModelService3000ServiceModelService3000
        Case wmiEnumClassWin32PerfRawDataSMSvcHost3000SMSvcHost3000
            SelectWmiClassName = WmiClassNameWin32PerfRawDataSMSvcHost3000SMSvcHost3000
        Case wmiEnumClassWin32PerfRawDataSMSvcHost4000SMSvcHost4000
            SelectWmiClassName = WmiClassNameWin32PerfRawDataSMSvcHost4000SMSvcHost4000
        Case wmiEnumClassWin32PerfRawDataSpoolerPrintQueue
            SelectWmiClassName = WmiClassNameWin32PerfRawDataSpoolerPrintQueue
        Case wmiEnumClassWin32PerfRawDataStorageStatsHyperVVirtualStorageDevice
            SelectWmiClassName = WmiClassNameWin32PerfRawDataStorageStatsHyperVVirtualStorageDevice
        Case wmiEnumClassWin32PerfRawDataTapiSrvTelephony
            SelectWmiClassName = WmiClassNameWin32PerfRawDataTapiSrvTelephony
        Case wmiEnumClassWin32PerfRawDataTBSTBScounters
            SelectWmiClassName = WmiClassNameWin32PerfRawDataTBSTBScounters
        Case wmiEnumClassWin32PerfRawDataTcpipICMP
            SelectWmiClassName = WmiClassNameWin32PerfRawDataTcpipICMP
        Case wmiEnumClassWin32PerfRawDataTcpipICMPv6
            SelectWmiClassName = WmiClassNameWin32PerfRawDataTcpipICMPv6
        Case wmiEnumClassWin32PerfRawDataTcpipIPv4
            SelectWmiClassName = WmiClassNameWin32PerfRawDataTcpipIPv4
        Case wmiEnumClassWin32PerfRawDataTcpipIPv6
            SelectWmiClassName = WmiClassNameWin32PerfRawDataTcpipIPv6
        Case wmiEnumClassWin32PerfRawDataTcpipNBTConnection
            SelectWmiClassName = WmiClassNameWin32PerfRawDataTcpipNBTConnection
        Case wmiEnumClassWin32PerfRawDataTcpipNetworkAdapter
            SelectWmiClassName = WmiClassNameWin32PerfRawDataTcpipNetworkAdapter
        Case wmiEnumClassWin32PerfRawDataTcpipNetworkInterface
            SelectWmiClassName = WmiClassNameWin32PerfRawDataTcpipNetworkInterface
        Case wmiEnumClassWin32PerfRawDataTcpipTCPv4
            SelectWmiClassName = WmiClassNameWin32PerfRawDataTcpipTCPv4
        Case wmiEnumClassWin32PerfRawDataTcpipTCPv6
            SelectWmiClassName = WmiClassNameWin32PerfRawDataTcpipTCPv6
        Case wmiEnumClassWin32PerfRawDataTcpipUDPv4
            SelectWmiClassName = WmiClassNameWin32PerfRawDataTcpipUDPv4
        Case wmiEnumClassWin32PerfRawDataTcpipUDPv6
            SelectWmiClassName = WmiClassNameWin32PerfRawDataTcpipUDPv6
        Case wmiEnumClassWin32PerfRawDataTCPIPCountersTCPIPPerformanceDiagnostics
            SelectWmiClassName = WmiClassNameWin32PerfRawDataTCPIPCountersTCPIPPerformanceDiagnostics
        Case wmiEnumClassWin32PerfRawDataTermServiceTerminalServicesSession
            SelectWmiClassName = WmiClassNameWin32PerfRawDataTermServiceTerminalServicesSession
        Case wmiEnumClassWin32PerfRawDataUGathererSearchGathererProjects
            SelectWmiClassName = WmiClassNameWin32PerfRawDataUGathererSearchGathererProjects
        Case wmiEnumClassWin32PerfRawDataUGTHRSVCSearchGatherer
            SelectWmiClassName = WmiClassNameWin32PerfRawDataUGTHRSVCSearchGatherer
        Case wmiEnumClassWin32PerfRawDatausbhubUSB
            SelectWmiClassName = WmiClassNameWin32PerfRawDatausbhubUSB
        Case wmiEnumClassWin32PerfRawDataVidPerfProviderHyperVVMVidNumaNode
            SelectWmiClassName = WmiClassNameWin32PerfRawDataVidPerfProviderHyperVVMVidNumaNode
        Case wmiEnumClassWin32PerfRawDataVidPerfProviderHyperVVMVidPartition
            SelectWmiClassName = WmiClassNameWin32PerfRawDataVidPerfProviderHyperVVMVidPartition
        Case wmiEnumClassWin32PerfRawDataVmbusStatsHyperVVirtualMachineBus
            SelectWmiClassName = WmiClassNameWin32PerfRawDataVmbusStatsHyperVVirtualMachineBus
        Case wmiEnumClassWin32PerfRawDataVmmsVirtualMachineStatsHyperVVirtualMachineHealthSummary
            SelectWmiClassName = WmiClassNameWin32PerfRawDataVmmsVirtualMachineStatsHyperVVirtualMachineHealthSummary
        Case wmiEnumClassWin32PerfRawDataVmmsVirtualMachineStatsHyperVVirtualMachineSummary
            SelectWmiClassName = WmiClassNameWin32PerfRawDataVmmsVirtualMachineStatsHyperVVirtualMachineSummary
        Case wmiEnumClassWin32PerfRawDataVmTaskManagerStatsHyperVTaskManagerDetail
            SelectWmiClassName = WmiClassNameWin32PerfRawDataVmTaskManagerStatsHyperVTaskManagerDetail
        Case wmiEnumClassWin32PerfRawDataW3SVCWebService
            SelectWmiClassName = WmiClassNameWin32PerfRawDataW3SVCWebService
        Case wmiEnumClassWin32PerfRawDataW3SVCWebServiceCache
            SelectWmiClassName = WmiClassNameWin32PerfRawDataW3SVCWebServiceCache
        Case wmiEnumClassWin32PerfRawDataW3SVCW3WPCounterProviderW3SVCW3WP
            SelectWmiClassName = WmiClassNameWin32PerfRawDataW3SVCW3WPCounterProviderW3SVCW3WP
        Case wmiEnumClassWin32PerfRawDataWASW3WPCounterProviderWASW3WP
            SelectWmiClassName = WmiClassNameWin32PerfRawDataWASW3WPCounterProviderWASW3WP
        Case wmiEnumClassWin32PerfRawDataWindowsMediaPlayerWindowsMediaPlayerMetadata
            SelectWmiClassName = WmiClassNameWin32PerfRawDataWindowsMediaPlayerWindowsMediaPlayerMetadata
        Case wmiEnumClassWin32PerfRawDataWindowsWorkflowFoundation3000WindowsWorkflowFoundation
            SelectWmiClassName = WmiClassNameWin32PerfRawDataWindowsWorkflowFoundation3000WindowsWorkflowFoundation
        Case wmiEnumClassWin32PerfRawDataWindowsWorkflowFoundation4000WFSystemWorkflow4000
            SelectWmiClassName = WmiClassNameWin32PerfRawDataWindowsWorkflowFoundation4000WFSystemWorkflow4000
        Case wmiEnumClassWin32PerfRawDataWorkflowServiceHost4000WorkflowServiceHost4000
            SelectWmiClassName = WmiClassNameWin32PerfRawDataWorkflowServiceHost4000WorkflowServiceHost4000
        Case wmiEnumClassWin32PerfRawDataWSearchIdxPiSearchIndexer
            SelectWmiClassName = WmiClassNameWin32PerfRawDataWSearchIdxPiSearchIndexer
        Case wmiEnumClassEventViewerConsumer
            SelectWmiClassName = WmiClassNameEventViewerConsumer
        Case wmiEnumClassNTEventlogProviderConfig
            SelectWmiClassName = WmiClassNameNTEventlogProviderConfig
        Case wmiEnumClassOfficeSoftwareProtectionProduct
            SelectWmiClassName = WmiClassNameOfficeSoftwareProtectionProduct
        Case wmiEnumClassOfficeSoftwareProtectionService
            SelectWmiClassName = WmiClassNameOfficeSoftwareProtectionService
        Case wmiEnumClassOfficeSoftwareProtectionTokenActivationLicense
            SelectWmiClassName = WmiClassNameOfficeSoftwareProtectionTokenActivationLicense
        Case wmiEnumClassRegistryEvent
            SelectWmiClassName = WmiClassNameRegistryEvent
        Case wmiEnumClassRegistryKeyChangeEvent
            SelectWmiClassName = WmiClassNameRegistryKeyChangeEvent
        Case wmiEnumClassRegistryTreeChangeEvent
            SelectWmiClassName = WmiClassNameRegistryTreeChangeEvent
        Case wmiEnumClassRegistryValueChangeEvent
            SelectWmiClassName = WmiClassNameRegistryValueChangeEvent
        Case wmiEnumClassScriptingStandardConsumerSetting
            SelectWmiClassName = WmiClassNameScriptingStandardConsumerSetting
        Case wmiEnumClassSoftwareLicensingProduct
            SelectWmiClassName = WmiClassNameSoftwareLicensingProduct
        Case wmiEnumClassSoftwareLicensingService
            SelectWmiClassName = WmiClassNameSoftwareLicensingService
        Case wmiEnumClassSoftwareLicensingTokenActivationLicense
            SelectWmiClassName = WmiClassNameSoftwareLicensingTokenActivationLicense
        Case wmiEnumClassStdRegProv
            SelectWmiClassName = WmiClassNameStdRegProv
    End Select
End Function

