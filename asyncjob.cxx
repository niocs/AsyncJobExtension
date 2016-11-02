#include <asyncjob.hxx>
#include <com/sun/star/task/XJobListener.hpp>
#include <com/sun/star/beans/NamedValue.hpp>
#include <com/sun/star/frame/DispatchResultEvent.hpp>
#include <com/sun/star/frame/DispatchResultState.hpp>
#include <com/sun/star/frame/XFrame.hpp>
#include <com/sun/star/frame/XController.hpp>
#include <com/sun/star/frame/XModel.hpp>
#include <com/sun/star/uno/XComponentContext.hpp>
#include <com/sun/star/sheet/XSpreadsheetDocument.hpp>
#include <com/sun/star/sheet/XSpreadsheets.hpp>
#include <com/sun/star/sheet/XSpreadsheet.hpp>
#include <com/sun/star/table/XCell.hpp>

#include <cppuhelper/supportsservice.hxx>
#include <rtl/ustring.hxx>
#include <rtl/ustrbuf.hxx>

#include <vector>

using rtl::OUString;
using rtl::OUStringBuffer;
using namespace com::sun::star::uno;
using namespace com::sun::star::frame;
using namespace com::sun::star::sheet;
using namespace com::sun::star::table;
using com::sun::star::beans::NamedValue;
using com::sun::star::task::XJobListener;
using com::sun::star::lang::IllegalArgumentException;
using com::sun::star::frame::DispatchResultEvent;
using com::sun::star::frame::DispatchResultEvent;


// This is the service name an Add-On has to implement
#define SERVICE_NAME "com.sun.star.task.AsyncJob"

// Helper functions for the implementation of UNO component interfaces.
OUString AsyncJobImpl_getImplementationName()
throw (RuntimeException)
{
    return OUString ( IMPLEMENTATION_NAME );
}

Sequence< OUString > SAL_CALL AsyncJobImpl_getSupportedServiceNames()
throw (RuntimeException)
{
    Sequence < OUString > aRet(1);
    OUString* pArray = aRet.getArray();
    pArray[0] =  OUString ( SERVICE_NAME );
    return aRet;
}

Reference< XInterface > SAL_CALL AsyncJobImpl_createInstance( const Reference< XComponentContext > & rContext)
    throw( Exception )
{
    return (cppu::OWeakObject*) new AsyncJobImpl( rContext );
}

// Implementation of the recommended/mandatory interfaces of a UNO component.
// XServiceInfo
OUString SAL_CALL AsyncJobImpl::getImplementationName()
    throw (RuntimeException)
{
    return AsyncJobImpl_getImplementationName();
}

sal_Bool SAL_CALL AsyncJobImpl::supportsService( const OUString& rServiceName )
    throw (RuntimeException)
{
    return cppu::supportsService(this, rServiceName);
}

Sequence< OUString > SAL_CALL AsyncJobImpl::getSupportedServiceNames(  )
    throw (RuntimeException)
{
    return AsyncJobImpl_getSupportedServiceNames();
}



// XAsyncJob method implementations

void SAL_CALL AsyncJobImpl::executeAsync( const Sequence<NamedValue>& rArgs,
					  const Reference<XJobListener>& rxListener )
    throw(IllegalArgumentException, RuntimeException)
{
    printf("DEBUG>>> Called executeAsync() : this = %p\n", this); fflush(stdout);
    
    AsyncJobImplInfo aJobInfo;
    OUString aErr = validateGetInfo( rArgs, rxListener, aJobInfo );
    if ( !aErr.isEmpty() )
    {
	sal_Int16 nArgPos = 0;	
	if ( aErr.startsWith( "Listener" ) )
	    nArgPos = 1;

	throw IllegalArgumentException(
	    aErr,
	    // resolve to XInterface reference:
	    static_cast< ::cppu::OWeakObject * >(this),
	    nArgPos ); // argument pos
    }
    
    writeInfoToSheet( aJobInfo );

    bool bIsDispatch = aJobInfo.aEnvType.equalsAscii("DISPATCH");
    Sequence<NamedValue> aReturn( ( bIsDispatch ? 1 : 0 ) );

    if ( bIsDispatch )
    {
	aReturn[0].Name  = "SendDispatchResult";
	DispatchResultEvent aResultEvent;
	aResultEvent.Source = (cppu::OWeakObject*)this;
	aResultEvent.State = DispatchResultState::SUCCESS;
	aResultEvent.Result <<= true;
	aReturn[0].Value <<= aResultEvent;
    }

    rxListener->jobFinished( Reference<com::sun::star::task::XAsyncJob>(this), makeAny(aReturn));
    
}


OUString AsyncJobImpl::validateGetInfo( const Sequence<NamedValue>& rArgs,
					const Reference<XJobListener>& rxListener,
					AsyncJobImpl::AsyncJobImplInfo& rJobInfo )
{
    if ( !rxListener.is() )
	return OUString("Listener : invalid listener");

    // Extract all sublists from rArgs.
    Sequence<NamedValue> aGenericConfig;
    Sequence<NamedValue> aJobConfig;
    Sequence<NamedValue> aEnvironment;
    Sequence<NamedValue> aDynamicData;

    sal_Int32 nNumNVs = rArgs.getLength();
    for ( sal_Int32 nIdx = 0; nIdx < nNumNVs; ++nIdx )
    {
	if ( rArgs[nIdx].Name.equalsAscii("Config") )
	{
	    rArgs[nIdx].Value >>= aGenericConfig;
	    formatOutArgs( aGenericConfig, "Config", rJobInfo.aGenericConfigList );
	}
	else if ( rArgs[nIdx].Name.equalsAscii("JobConfig") )
	{
	    rArgs[nIdx].Value >>= aJobConfig;
	    formatOutArgs( aJobConfig, "JobConfig", rJobInfo.aJobConfigList );
	}
	else if ( rArgs[nIdx].Name.equalsAscii("Environment") )
	{
	    rArgs[nIdx].Value >>= aEnvironment;
	    formatOutArgs( aEnvironment, "Environment", rJobInfo.aEnvironmentList );
	}
	else if ( rArgs[nIdx].Name.equalsAscii("DynamicData") )
	{
	    rArgs[nIdx].Value >>= aDynamicData;
	    formatOutArgs( aDynamicData, "DynamicData", rJobInfo.aDynamicDataList );
	}
    }

    // Analyze the environment info. This sub list is the only guaranteed one!
    if ( !aEnvironment.hasElements() )
	return OUString("Args : no environment");

    sal_Int32 nNumEnvEntries = aEnvironment.getLength();
    for ( sal_Int32 nIdx = 0; nIdx < nNumEnvEntries; ++nIdx )
    {
	if ( aEnvironment[nIdx].Name.equalsAscii("EnvType") )
	    aEnvironment[nIdx].Value >>= rJobInfo.aEnvType;
	
	else if ( aEnvironment[nIdx].Name.equalsAscii("EventName") )
	    aEnvironment[nIdx].Value >>= rJobInfo.aEventName;

	else if ( aEnvironment[nIdx].Name.equalsAscii("Frame") )
	    aEnvironment[nIdx].Value >>= rJobInfo.xFrame;
    }

    // Further the environment property "EnvType" is required as minimum.

    if ( rJobInfo.aEnvType.isEmpty() ||
	 ( ( !rJobInfo.aEnvType.equalsAscii("EXECUTOR") ) &&
	   ( !rJobInfo.aEnvType.equalsAscii("DISPATCH") )
	     )	)
	return OUString("Args : \"" + rJobInfo.aEnvType + "\" isn't a valid value for EnvType");

    // Analyze the set of shared config data.
    if ( aGenericConfig.hasElements() )
    {
	sal_Int32 nNumGenCfgEntries = aGenericConfig.getLength();
	for ( sal_Int32 nIdx = 0; nIdx < nNumGenCfgEntries; ++nIdx )
	    if ( aGenericConfig[nIdx].Name.equalsAscii("Alias") )
		aGenericConfig[nIdx].Value >>= rJobInfo.aAlias;
    }

    return OUString("");
}


void AsyncJobImpl::formatOutArgs( const Sequence<NamedValue>& rList,
				  const OUString aListName,
				  OUString& rFormattedList )
{
    OUStringBuffer aBuf( 512 );
    aBuf.append("List \"" + aListName + "\" : ");
    if ( !rList.hasElements() )
	aBuf.append("0 items");
    else
    {
	sal_Int32 nNumEntries = rList.getLength();
	aBuf.append(nNumEntries).append(" item(s) : { ");
	for ( sal_Int32 nIdx = 0; nIdx < nNumEntries; ++nIdx )
	{
	    OUString aVal;
	    rList[nIdx].Value >>= aVal;
	    aBuf.append("\"" + rList[nIdx].Name + "\" : \"" + aVal + "\", ");
	}
	aBuf.append(" }");
    }

    rFormattedList = aBuf.toString();
}


void WriteStringsToSheet( const Reference< XFrame > &rxFrame, const std::vector<OUString>& rStrings )
{
    if ( !rxFrame.is() )
	return;
    
    Reference< XController > xCtrl = rxFrame->getController();
    if ( !xCtrl.is() )
	return;

    Reference< XModel > xModel = xCtrl->getModel();
    if ( !xModel.is() )
	return;

    Reference< XSpreadsheetDocument > xSpreadsheetDocument( xModel, UNO_QUERY );
    if ( !xSpreadsheetDocument.is() )
	return;

    Reference< XSpreadsheets > xSpreadsheets = xSpreadsheetDocument->getSheets();
    if ( !xSpreadsheets->hasByName("Sheet1") )
	return;

    Any aSheet = xSpreadsheets->getByName("Sheet1");

    Reference< XSpreadsheet > xSpreadsheet( aSheet, UNO_QUERY );

    size_t nNumStrings = rStrings.size();
    for ( size_t nIdx = 0; nIdx < nNumStrings; ++nIdx )
    {
	Reference< XCell > xCell = xSpreadsheet->getCellByPosition(0, nIdx);
	xCell->setFormula(rStrings[nIdx]);
	printf("DEBUG>>> Wrote \"%s\" to Cell A%d\n",
	       OUStringToOString( rStrings[nIdx], RTL_TEXTENCODING_ASCII_US ).getStr(), nIdx); fflush(stdout);
    }
}

void AsyncJobImpl::writeInfoToSheet( const AsyncJobImpl::AsyncJobImplInfo& rJobInfo )
{
    if ( !rJobInfo.xFrame.is() )
    {
	printf("DEBUG>>> Frame passed is null, cannot write to sheet !\n");
	fflush(stdout);
    }
    std::vector<OUString> aStrList = {
	"EnvType = " + rJobInfo.aEnvType,
	"EventName = " + rJobInfo.aEventName,
	"Alias = " + rJobInfo.aAlias,
	rJobInfo.aGenericConfigList,
	rJobInfo.aJobConfigList,
	rJobInfo.aEnvironmentList,
	rJobInfo.aDynamicDataList
    };

    WriteStringsToSheet( rJobInfo.xFrame, aStrList );
}

