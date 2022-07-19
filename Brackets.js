var _ = {
	GatherArguments : function (objFunction) {
		return function () {
			objFunction.Apply(arguments, arguments.callee);
			return objFunction.ReturnValue;
		};
	},
	SpreadArguments : function (objFunction, arrArguments){
		return objFunction.apply(null, arrArguments.toArray());
	}
}
