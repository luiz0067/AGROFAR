String.prototype.replaceAll = function(pcFrom, pcTo){	
	var i = this.indexOf(pcFrom);	
	var c = this;		
	while (i > -1){		
		c = c.replace(pcFrom, pcTo); 		
		i = c.indexOf(pcFrom);	
	}	
	return c;
};	
String.prototype.trim = function() { 
	return this.replace(/^\s+|\s+$/, ''); 
};

String.prototype.count=function( sWhat ) {
	return this.split( sWhat ).length - 1;
}
