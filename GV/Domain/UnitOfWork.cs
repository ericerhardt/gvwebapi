using System;
using System.Data;
using GV.CoFreedomDomain;
using NHibernate;

namespace GV.Domain
{
    public interface IUnitOfWork : IDisposable
    {
        ISession CurrentSession { get; }
        void Commit();
        void Rollback();
    }

      public class UnitOfWork : ICoFreedomUnitOfWork
    {
        private readonly ISessionFactoryHelper _sessionFactoryHelper;
        private ISession _currentSession;
        private bool _disposed;
        private IsolationLevel _isolationLevel = IsolationLevel.ReadCommitted;

        public UnitOfWork(ISessionFactoryHelper sessionFactoryHelper)
        {
            _sessionFactoryHelper = sessionFactoryHelper;
        }

        ~UnitOfWork()
        {
            Dispose(false);
        }

        public void SetIsolationLevel(IsolationLevel isolationLevel)
        {
            _isolationLevel = isolationLevel;
        }

        public ISession CurrentSession
        {
            get
            {
                if (_currentSession == null  || _currentSession.Connection.State != ConnectionState.Open || _currentSession.IsOpen == false || _currentSession.IsConnected == false)
                    _currentSession = OpenSession();
                return _currentSession;
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        public void Commit()
        {
            CheckDisposed();

            try
            {
                if (_currentSession != null && _currentSession.GetSessionImplementation().ConnectionManager.IsInActiveTransaction)
                    _currentSession.Transaction.Commit();

                _currentSession?.BeginTransaction(_isolationLevel);
            }
            catch (Exception)
            {
                RollbackAndCloseSession(true);
                throw;
            }
        }

        public void Rollback()
        {
            CheckDisposed();
            RollbackAndCloseSession();
        }

        private ISession OpenSession()
        {
            CheckDisposed();

            //Prints NHibernate SQL to output window when in debug mode.  Based on post found here: https://stackoverflow.com/questions/129133/how-do-i-view-the-sql-that-is-generated-by-nhibernate
            #if LOCAL
                var session = _sessionFactoryHelper.CurrentSessionFactory.OpenSession(new SQLDebugOutput());            
            #else
                var session = _sessionFactoryHelper.CurrentSessionFactory.OpenSession();
            #endif

            session.FlushMode = FlushMode.Commit;
            session.BeginTransaction(_isolationLevel);
            return session;
        }

        private void Dispose(bool disposing)
        {
            //            Debug($"Dispose: {disposing}");

            if (_disposed) return;

            if (disposing)
                RollbackAndCloseSession(true);

            _currentSession = null;

            _disposed = true;
        }

        private void RollbackAndCloseSession(bool ignoreException = false)
        {
            try
            {
                if (_currentSession != null &&
                    _currentSession.GetSessionImplementation().ConnectionManager.IsInActiveTransaction)
                    _currentSession.Transaction.Rollback();
            }
            catch (Exception ex)
            {
                if (!ignoreException)
                    throw new Exception("Error in RollbackAndCloseSession", ex);
            }
            finally
            {
                if (_currentSession != null)
                {
                    _currentSession.Dispose();
                    _currentSession = null;
                }
            }
        }

        private void CheckDisposed()
        {
            if (_disposed) throw new ObjectDisposedException(GetType().FullName);
        }

    }
}