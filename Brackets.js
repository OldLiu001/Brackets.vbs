var _ = {
	WrapFunction : function (objAnonymousFunction) {
		return function (objArguments) {
			objAnonymousFunction.Apply(objArguments);
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
