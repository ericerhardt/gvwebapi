using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using GV.CoFreedomDomain;
using NHibernate.Linq;
using NHibernate.Transform;

namespace GV.Domain
{
    public interface IRepository
    {
        T Load<T>(long id);
        T Load<T>(long? id);
        T Load<T>(int id);

        T Get<T>(long id);
        T Get<T>(int id);
        
        T Get<T>(Expression<Func<T, bool>> predicate);

        IQueryable<T> Find<T>();

        T Add<T>(T entity);
        T Remove<T>(T entity);
        void Remove<T>(long id);
        IList<T> ExecuteSQL<T>(string query);
    }

    public class Repository : ICoFreedomRepository
    {
          private readonly IUnitOfWork _unitOfWork;

        public Repository(IUnitOfWork unitOfWork)
        {
            _unitOfWork = unitOfWork;
        }

        public T Load<T>(long id)
        {
            return id == 0L ? default(T) : _unitOfWork.CurrentSession.Load<T>(id);
        }
        
        public T Load<T>(long? id)
        {
            return id.HasValue ? Load<T>(id.Value) : default(T);
        }

        public T Load<T>(int id)
        {
            return id == 0 ? default(T) : _unitOfWork.CurrentSession.Load<T>(id);
        }
         
        public T Get<T>(long id)
        {
            return id == 0L ? default(T) : _unitOfWork.CurrentSession.Get<T>(id);
        }

        public T Get<T>(int id)
        {
            return id == 0 ? default(T) : _unitOfWork.CurrentSession.Get<T>(id);
        }
        
        public T Get<T>(Expression<Func<T, bool>> predicate)
        {
            return Find<T>().SingleOrDefault(predicate);
        }

        public IQueryable<T> Find<T>()
        {
            return _unitOfWork.CurrentSession.Query<T>();
        }

        public T Add<T>(T entity)
        {
            _unitOfWork.CurrentSession.Save(entity);
            return entity;
        }

        public T Remove<T>(T entity)
        {
            _unitOfWork.CurrentSession.Delete(entity);
            return entity;
        }

        public void Remove<T>(long id)
        {
            Remove(Load<T>(id));
        }
        
        public IList<T> ExecuteSQL<T>(string query)
        {
            return _unitOfWork.CurrentSession.CreateSQLQuery(query).SetResultTransformer(Transformers.AliasToBean<T>()).List<T>();
        }
    }
}