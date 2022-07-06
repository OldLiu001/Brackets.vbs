function WrapFunction(objLambda) {
	return function (objArguments) {
		objLambda.Apply(objArguments);
		return objLambda.ReturnValue;
	}
}

function WrapArguments(objLambda) {
	return function () {
		return objLambda(arguments);
	}
}

//WScript.Echo(Brackets.If(true,1,2));

// const forEach = (array,fn)=>{
// 	for(let i=0;i<array.length;i++){
// 	  fn(array[i])
// 	}
//   }
  
//   const unless = (predicate,fn)=>{
// 	  if(!predicate)
// 		  fn()
//   }
  
//   /*
//   times_ = function(time_,fn){
// 	  for(let i=0;i<time_;i++){
// 		  fn(i);
// 	  }
//   }*/
//   const every = (arr,fn)=>{
// 	  let result = true;
// 	  for(let i =0;i<arr.length;i++){
// 		  result = result&&fn(arr[i])
// 	  }
// 	  return true;
//   }
//   const every2 = (arr,fn)=>{
// 	  let result = true;
// 	  for(const element of arr){
// 		  result = result&&fn(element)
// 	  }
// 	  return true;
//   }
  
//   const some = (arr,fn)=>{
// 	  let result = false;
// 	  for(const element of arr){
// 		  result = result||fn(element)
// 	  }
// 	  return true;
//   }
  
//   /*
//   const sortBy = (property)=>{
// 	  return (a,b) => {
// 		  return (a[property]<b[property])?-1:(a[property]>b[property])?1:0
// 	  }
//   }
  
//   const once = (fn)=>{
// 	  let done = false;
// 	  return function(){
// 		  return done?undefined:((done=true),fn.apply(this,arguments))
// 	  }
//   }
  
//   const memoized = (fn) => {
// 	  const lookupTable = {};
// 	  return (arg) => lookupTable[arg] || (lookupTable[arg]=fn(arg));
//   }
  
//   */
  
//   /*
//   const map = (array,fn) => {
// 	  let results= [];
// 	  for(const value of array)
// 		  results.push(fn(value))
// 	  return results;
//   }
//   */
//   const map = (array,fn) => {
// 	  let results= [];
// 	  for(let i=0;i<array.length;i++){
// 	  results.push(fn(array[i]))
// 		}
// 	  return results;
//   }
  
//   /*
//   const filter = (array,fn) => {
// 	  let results= [];
// 	  for(const value of array)
// 		  fn(value) ? results.push(value) : undefined
// 	  return results;
//   }
//   */
//   const filter = (array,fn) => {
// 	  let results= [];
// 	  for(let i=0;i<array.length;i++){
// 		  if(fn(array[i])) 
// 		  results.push(array[i]) 
// 	  }
// 	  return results;
//   }
  
//   const concatAll = (array) => {
// 	  let results = [];
// 	  for(const value of array){
// 		  results.push.apply(results,value)  //重点！！
// 	  }
// 	  return results;
//   }
//   function flatten(array){
// 	  var result = [];
// 	  var toStr = Object.prototype.toString;
// 	  for(var i=0;i<array.length;i++){
// 		  var element = array[i];
// 		  if(toStr.call(element) === "[object Array]"){ //Array.isArray(element) === true
// 			  result = result.concat(flatten(element)); //[...result,...flatten(element)]
// 		  }
// 		  else{
// 			  result.push(element);
// 		  }
// 	  }
// 	  return result;
//   }
  
//   const reduce = (array,fn)=>{
// 	  if(array.length >= 1) {
// 	  let accumlator = array[0];
// 	  for(var i=1;i<array.length;i++){
// 		  accumlator = fn(accumlator,value);
// 	  }
// 	  return [accumlator]
// 	  }else return [];
//   }
//   /*
//   const reduce = (array,fn,initialValue)=>{
// 	  let accumlator;
// 	  if(initialValue != undefined)
// 		  accumlator = initialValue;
// 	  else
// 		  accumlator = array[0];
// 	  //当initialValue未定义时，我们需要从第二个元素开始循环数组
// 	  if(initialValue === undefined){
// 		  for(let i=1; i<array.length;i++){
// 			  accumlator = fn(accumlator,array[i])
// 		  }
// 	  }else{//如果initialValue由调用者传入，我们就需要遍历整个数组。
// 		  for(const value of array){
// 			  accumlator = fn(accumlator,value);
// 		  }
// 	  }
// 	  return [accumlator]
//   }*/
  
//   const zip = (leftArr,rightArr,fn) => {
// 	  let index,results=[];
// 	  for(index=0;index<Math.min(leftArr.length,rightArr.length);index++){
// 		  results.push(fn(leftArr[index],rightArr[index]));
// 	  }
// 	  return results;
//   }
  
//   /*
//   let curry = (fn) => {
// 	  if(typeof fn!=='function'){
// 		  throw Error('No function provided')
// 	  }
// 	  return function curriedFn(){ //返回函数是一个变参函数
// 		  return fn(...arguments) 
// 	  }
// 	  //采用如下写法也ok
// 	  // return function curriedFn(...args){
// 	  //  return fn(...args) 
// 	  // }
//   }
//   let curry = (fn) => {
// 	  if(typeof fn!=='function'){
// 		  throw Error('No function provided')
// 	  }
// 	  return function curriedFn(...args){ //args是一个数组
// 		  if(args.length < fn.length){ //检查...args传入的参数长度是否小于函数参数列表的长度
// 			  return function(){
// 				  args = [...args,...arguments]
// 				  return curriedFn(...args)
// 			  };
// 		  }
// 		  return fn(...args) //不小于，就和之前一样调用整个函数
// 	  }
//   }
//   const partial = function(fn, ...partialArgs){
// 	  let args = partialArgs;
// 	  return function(...fullArguments){
// 		  let arg = 0;
// 		  for(let i=0;i<args.length && arg<fullArguments.length;i++){
// 			  if(args[i]===undefined){
// 				  args[i] = fullArguments[arg++];
// 			  }
// 		  }
// 		  return fn.apply(null,args)
// 	  }
//   }
  
//   柯里化
//   const _end = Symbol.for('All arguments ready');
//   let currying = function(fn, len = fn.length) {
// 	return function f(...args) {
// 	  if(args.length === 0) args.push(undefined);
// 	  if(args[args.length - 1] === _end) {
// 		return fn.apply(this, args.slice(0, -1));
// 	  }
// 	  if(args.length >= len) {
// 		return fn.apply(this, args);
// 	  }
// 	  return currying(fn.bind(this, ...args), len - args.length);
// 	}
//   }
//   currying = currying(currying, 2);
  
//   compose
//   partial
//   管道/序列（pipeline/sequence）：从左至右处理数据流的过程称为管道/序列
//   函子 純函數的方法進行錯誤處理 是一種容器
//   Container 容器
//   Generator 生成器/迭代器
//   */
  
