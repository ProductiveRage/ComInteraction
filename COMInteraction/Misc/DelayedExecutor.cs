using System;

namespace COMInteraction.Misc
{
	/// <summary>
	/// This is similar to the .Net 4's Lazy class with the isThreadSafe argument set to true
	/// </summary>
	public class DelayedExecutor<T> where T : class
	{
		private readonly Func<T> _work;
		private readonly object _lock;
		private volatile Result _result;
		public DelayedExecutor(Func<T> work)
		{
			if (work == null)
				throw new ArgumentNullException("work");

			_work = work;
			_lock = new object();
			_result = null;
		}

		public T Value
		{
			get
			{
				if (_result == null)
				{
					lock (_lock)
					{
						if (_result == null)
						{
							try
							{
								_result = Result.Success(_work());
							}
							catch(Exception e)
							{
								_result = Result.Failulre(e);
							}
						}
					}
				}
				if (_result.Error != null)
					throw _result.Error;
				return _result.Value;
			}
		}

		private class Result
		{
			public static Result Success(T value)
			{
				return new Result(value, null);
			}
			public static Result Failulre(Exception error)
			{
				if (error == null)
					throw new ArgumentNullException("error");

				return new Result(null, error);
			}
			private Result(T value, Exception error)
			{
				Value = value;
			}

			public T Value { get; private set; }

			public Exception Error { get; private set; }
		}
	}
}
