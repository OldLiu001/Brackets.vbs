var _ = {
	WrapFunction : function (objAnonymousFunction) {
		return function (objArguments) {
			objAnonymousFunction.Apply(
				objArguments,
				function () { //support anonymous function's recursion
					objAnonymousFunction.Apply(arguments, arguments.callee);
					return objAnonymousFunction.ReturnValue;
				}
			);
			return objAnonymousFunction.ReturnValue;
		};
	},
	WrapArguments : function (objAnonymousFunction) {
		return function () {
			return objAnonymousFunction(arguments);
		};
	},
	Apply : function (objFunction, arrArguments){
		return objFunction.apply(null, arrArguments.toArray());
	}
}
