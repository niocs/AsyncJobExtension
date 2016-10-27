#ifndef INCO_NIOCS_TEST_ASYNCJOBEXTENSION_ADDON_HXX
#define INCO_NIOCS_TEST_ASYNCJOBEXTENSION_ADDON_HXX

#include <com/sun/star/task/XAsyncJob.hpp>
#include <com/sun/star/lang/XServiceInfo.hpp>
#include <com/sun/star/lang/IllegalArgumentException.hpp>
#include <cppuhelper/implbase2.hxx>

#define IMPLEMENTATION_NAME "inco.niocs.test.AsyncJobImpl"

namespace com
{
    namespace sun
    {
        namespace star
        {
	    namespace frame
	    {
		class XFrame;
	    }
            namespace uno
            {
                class XComponentContext;
            }
	    namespace beans
	    {
		struct NamedValue;
	    }
	    namespace task
	    {
		class XJobListener;
	    }
        }
    }
}

class AsyncJobImpl : public cppu::WeakImplHelper2
<
    com::sun::star::task::XAsyncJob,
    com::sun::star::lang::XServiceInfo
>
{
    
private:
    ::com::sun::star::uno::Reference< ::com::sun::star::uno::XComponentContext > mxContext;

public:

    AsyncJobImpl( const ::com::sun::star::uno::Reference< ::com::sun::star::uno::XComponentContext > &rxContext )
        : mxContext( rxContext )
    {
	printf("DEBUG>>> Created AsyncJobImpl object : %p\n", this); fflush(stdout);
    }

    // XAsyncJob methods
    virtual void SAL_CALL executeAsync( const ::com::sun::star::uno::Sequence< ::com::sun::star::beans::NamedValue >& rArgs,
					const ::com::sun::star::uno::Reference< ::com::sun::star::task::XJobListener >& rxListener )
	throw(::com::sun::star::lang::IllegalArgumentException,
	      ::com::sun::star::uno::RuntimeException) override;

    // XServiceInfo methods
    virtual ::rtl::OUString SAL_CALL getImplementationName()
        throw (::com::sun::star::uno::RuntimeException);
    virtual sal_Bool SAL_CALL supportsService( const ::rtl::OUString& aServiceName )
        throw (::com::sun::star::uno::RuntimeException);
    virtual ::com::sun::star::uno::Sequence< ::rtl::OUString > SAL_CALL getSupportedServiceNames()
        throw (::com::sun::star::uno::RuntimeException);

    // A struct to store some job related info when executeAsync() is called
    struct AsyncJobImplInfo
    {
	::rtl::OUString aEnvType;
	::rtl::OUString aEventName;
	::rtl::OUString aAlias;
	::com::sun::star::uno::Reference< ::com::sun::star::frame::XFrame > xFrame;

	::rtl::OUString aGenericConfigList, aJobConfigList, aEnvironmentList, aDynamicDataList;
    };

private:

    ::rtl::OUString validateGetInfo( const ::com::sun::star::uno::Sequence< ::com::sun::star::beans::NamedValue >& rArgs,
				     const ::com::sun::star::uno::Reference< ::com::sun::star::task::XJobListener >& rxListener,
				     AsyncJobImplInfo& rJobInfo );
    
    void formatOutArgs( const ::com::sun::star::uno::Sequence< ::com::sun::star::beans::NamedValue >& rList,
			const ::rtl::OUString aListName,
			::rtl::OUString& rFormattedList );

    void writeInfoToSheet( const AsyncJobImplInfo& rJobInfo );
};

::rtl::OUString AsyncJobImpl_getImplementationName()
    throw ( ::com::sun::star::uno::RuntimeException );

sal_Bool SAL_CALL AsyncJobImpl_supportsService( const ::rtl::OUString& ServiceName )
    throw ( ::com::sun::star::uno::RuntimeException );

::com::sun::star::uno::Sequence< ::rtl::OUString > SAL_CALL AsyncJobImpl_getSupportedServiceNames()
    throw ( ::com::sun::star::uno::RuntimeException );

::com::sun::star::uno::Reference< ::com::sun::star::uno::XInterface >
SAL_CALL AsyncJobImpl_createInstance( const ::com::sun::star::uno::Reference< ::com::sun::star::uno::XComponentContext > & rContext)
    throw ( ::com::sun::star::uno::Exception );

#endif
