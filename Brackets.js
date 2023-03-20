var _ = {
	Expr : function (code) {
		return _.Unwrap(VbsExpr(code));
	},

	Unwrap : function (pkg) {
		return new VBArray(pkg).toArray()[0];
	},

	GatherArguments : function (objFunction) {
		return function () {
			objFunction.Apply(arguments, arguments.callee);
			return objFunction.ReturnValue;
		};
	},
	
	SpreadArguments : function (objFunction, arrArguments){
		return objFunction.apply(null, arrArguments.toArray());
	},
	
	If : function (cond, expr1, expr2) {
		// No short-circuit, all arguments will be evaluated.
		return cond ? expr1 : expr2;
	},
	
	Function : function (para, code, sbinds, abinds) {
		WScript.Echo(233);
		WScript.Echo(_.Expr("1+2"));
	}
}
zz = _;