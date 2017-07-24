Imports System
Imports System.Configuration
Imports System.Data
Imports System.IO
Imports System.Web

Namespace PogoloDataProvider.Data

    Public Enum CoreCacheType
        Host = 1
        Portal = 2
        Tab = 3
    End Enum

    Public Class DataCache

        Public Shared Function GetCache(ByVal CacheKey As String) As Object

            Dim objCache As System.Web.Caching.Cache = HttpRuntime.Cache

            Return objCache(CacheKey)

        End Function

        Public Shared Sub SetCache(ByVal CacheKey As String, ByVal objObject As Object)

            Dim objCache As System.Web.Caching.Cache = HttpRuntime.Cache

            objCache.Insert(CacheKey, objObject)

        End Sub

        Public Shared Sub SetCache(ByVal CacheKey As String, ByVal objObject As Object, ByVal objDependency As System.Web.Caching.CacheDependency)

            Dim objCache As System.Web.Caching.Cache = HttpRuntime.Cache

            objCache.Insert(CacheKey, objObject, objDependency)

        End Sub

        Public Shared Sub SetCache(ByVal CacheKey As String, ByVal objObject As Object, ByVal SlidingExpiration As Integer)

            Dim objCache As System.Web.Caching.Cache = HttpRuntime.Cache

            objCache.Insert(CacheKey, objObject, Nothing, DateTime.MaxValue, TimeSpan.FromSeconds(SlidingExpiration))

        End Sub

        Public Shared Sub SetCache(ByVal CacheKey As String, ByVal objObject As Object, ByVal AbsoluteExpiration As Date)

            Dim objCache As System.Web.Caching.Cache = HttpRuntime.Cache

            objCache.Insert(CacheKey, objObject, Nothing, AbsoluteExpiration, TimeSpan.Zero)

        End Sub

        Public Shared Sub RemoveCache(ByVal CacheKey As String)

            Dim objCache As System.Web.Caching.Cache = HttpRuntime.Cache

            If Not objCache(CacheKey) Is Nothing Then
                objCache.Remove(CacheKey)
            End If

        End Sub

    End Class

End Namespace