function WrapFunction(objAnonymousFunction) {
	return function (objArguments) {
		objAnonymousFunction.Apply(objArguments);
		return objAnonymousFunction.ReturnValue;
	};
}

function WrapArguments(objAnonymousFunction) {
	return function () {
		return objAnonymousFunction(arguments);
	};
}

function Apply(objFunction, arrArguments){
	return objFunction.apply(null, arrArguments.toArray());
}