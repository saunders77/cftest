/*
Office.initialize = function(reason){
    var debug = [];
    var debugUpdate = function(data){};
    function write(myText){
        debug.push([myText]);
        debugUpdate(debug);
    }

    function myDebug(setResult){
        debugUpdate = setResult;
    }
        

    function secondHighestTemp(temperatures){ 
        var highest = -273, secondHighest = -273;
        for(var i = 0; i < temperatures.length;i++){
            for(var j = 0; j < temperatures[i].length;j++){
                if(temperatures[i][j] >= highest){
                    secondHighest = highest;
                    highest = temperatures[i][j];
                }
                else if(temperatures[i][j] >= secondHighest){
                    secondHighest = temperatures[i][j];
                }
            }
        }
        return secondHighest;
    }

    Excel.Script.CustomFunctions = {};
    Excel.Script.CustomFunctions["CONTOSO"] = {};
    
    Excel.Script.CustomFunctions["CONTOSO"]["SECONDHIGHESTTEMP"] = {
        call: secondHighestTemp,
        description: "Returns the second highest from a range of temperatures",
        helpUrl: "https://example.com/help.html",
        result: {
            resultType: Excel.CustomFunctionValueType.number,
            resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
        },
        parameters: [
            {
                name: "temps",
                description: "the temperatures to be compared",
                valueType: Excel.CustomFunctionValueType.number,
                valueDimensionality: Excel.CustomFunctionDimensionality.matrix,
            },
        ],
        options: {
            batch: false,
            stream: false,
        }
    };
    
    Excel.run(function (context) {
        context.workbook.customFunctions.addAll();
        return context.sync().then(function(){

        });
    
    }).catch(function(error){
        console.log("error" + error);
    });
};

*/

Office.initialize = function(reason){
    
    //Office.context.ui.displayDialogAsync("https://www.michael-saunders.com/stocksapp/pages/info.html",{ height: 50, width: 50, displayInIframe: true }, function(){});

    //define the Contoso prefix
    Excel.Script.CustomFunctions = {};
    Excel.Script.CustomFunctions["CONTOSO"] = {};


    // sample synchronous function
    function add42 (a) {
        return a + 42;
    }    
    Excel.Script.CustomFunctions["CONTOSO"]["ADD42"] = {
        call: add42,
        description: "Finds the sum of two numbers and 42",
        helpUrl: "https://www.contoso.com/help.html",
        result: {
            resultType: Excel.CustomFunctionValueType.number,
            resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
        },
        parameters: [
            {
                name: "num",
                description: "the number",
                valueType: Excel.CustomFunctionValueType.number,
                valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
            },
        ],
        options:{ batch: false, stream: false }
    };


    // helper code for getting temperature
    var temps = {};
    temps["boiler"] = 104.3;
    temps["mixer"] = 44.0;
    temps["furnace"] = 586.9;
    furnaceHistory = [];
    function startTime(){
        temps["boiler"] += Math.pow(Math.random() - 0.45, 3) * 2;
        temps["mixer"] += Math.pow(Math.random() - 0.55, 3) * 2;
        temps["furnace"] += Math.pow(Math.random() - 0.40, 3) * 2;
        furnaceHistory.push([temps["furnace"]]);
        if(furnaceHistory.length > 50){
            furnaceHistory.shift();
        }
        setTimeout(startTime, 500);
    }
    startTime();
    function getTempFromServer(thermometerID, callback){
        setTimeout(function(){
            var data = {};
            data.temperature = temps[thermometerID].toFixed(1);
            callback(data);
        }, 200);
    }

    // demo functions

    function getTemperature(thermometerID){ 

        return new OfficeExtension.Promise(function(setResult, setError){ 
            
            
            getTempFromServer(thermometerID, function(data){ 
                setResult(data.temperature); 
            }); 
            
        }); 
        
    }

    function streamTemperature(thermometerID, interval, call){     
        if(thermometerID == "furnace"){
            temps["furnace"] = 630.2;
        }
        function getNextTemperature(){ 
            getTempFromServer(thermometerID, function(data){ 
                call.setResult(data.temperature); 
            }); 
            setTimeout(getNextTemperature, interval); 
        } 
        getNextTemperature(); 
    } 

    function secondHighestTemp(temperatures){ 
        var highest = -273, secondHighest = -273;
        for(var i = 0; i < temperatures.length;i++){
            for(var j = 0; j < temperatures[i].length;j++){
                if(temperatures[i][j] >= highest){
                    secondHighest = highest;
                    highest = temperatures[i][j];
                }
                else if(temperatures[i][j] >= secondHighest){
                    secondHighest = temperatures[i][j];
                }
            }
        }
        return secondHighest;
    }

    function trackTemperature(thermometerID, call){
        var output = [];
        
        for(var i = 0; i < 50; i++) output.push([0]);  
        if(thermometerID == "furnace"){
            output = furnaceHistory;
        } 
        function recordNextTemperature(){
            getTempFromServer(thermometerID, function(data){
                output.push([data.temperature]);
                output.shift();
                call.setResult(output);
            });
            setTimeout(recordNextTemperature, 500);
        }
        recordNextTemperature();
    } 

    Excel.Script.CustomFunctions["CONTOSO"]["TRACKTEMPERATURE"] = {
        call: trackTemperature,
        description: "Streams 25 seconds of temperature history",
        helpUrl: "https://example.com/help.html",
        result: {
            resultType: Excel.CustomFunctionValueType.number,
            resultDimensionality: Excel.CustomFunctionDimensionality.matrix,
        },
        parameters: [
            {
                name: "thermometer ID",
                description: "The thermometer to be measured",
                valueType: Excel.CustomFunctionValueType.string,
                valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
            },
        ],
        options: {
            batch: false,
            stream: true,
        }
    };

    /*
    // sample asynchronous function
    function getTemperature(thermometerID){
        return new OfficeExtension.Promise(function(setResult, setError){
            sendWebRequestExample(thermometerID, function(data){
                setResult(data.temperature);
            });
        });
    }
    */
    Excel.Script.CustomFunctions["CONTOSO"]["GETTEMPERATURE"] = {
        call: getTemperature,
        description: "Returns the temperature of a sensor",
        helpUrl: "https://www.contoso.com/help.html",
        result: {
            resultType: Excel.CustomFunctionValueType.number,
            resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
        },
        parameters: [
            {
                name: "thermometer ID",
                description: "The thermometer to be measured",
                valueType: Excel.CustomFunctionValueType.string,
                valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
            },
        ],
        options: { batch: false,  stream: false }
    };
/*
    // sample streaming function
    function incrementValue(increment, setResult){    
    	var result = 0;
        setInterval(function(){
            result += increment;
            setResult(result);
        }, 1000);
    }
    Excel.Script.CustomFunctions["CONTOSO"]["INCREMENTVALUE"] = {
        call: incrementValue,
        description: "Counts up from zero",
        helpUrl: "https://www.contoso.com/help.html",
        result: {
            resultType: Excel.CustomFunctionValueType.number,
            resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
        },
        parameters: [
            {
                name: "period",
                description: "the time between updates, in milliseconds",
                valueType: Excel.CustomFunctionValueType.number,
                valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
            },
        ],
        options: { batch: false,  stream: true }
    };
    
    // sample function that uses global variables to save state while streaming data
    var savedTemperatures = {};
    function refreshTemperature(thermometerID){
        sendWebRequestExample(thermometerID, function(data){
            savedTemperatures[thermometerID] = data.temperature;
        });
        setTimeout(function(){
            refreshTemperature(thermometerID);
        }, 1000);
    }
    function streamTemperature(thermometerID, setResult){    
        if(!savedTemperatures[thermometerID]){
            refreshTemperatures(thermometerID);
        }
        function getNextTemperature(){
            setResult(savedTemperatures[thermometerID]);
            setTimeout(getNextTemperature, 1000);
        }
        getNextTemperature();
    }
    */
    Excel.Script.CustomFunctions["CONTOSO"]["STREAMTEMPERATURE"] = {
        call: streamTemperature,
        description: "Returns the temperature of a sensor every second",
        helpUrl: "https://www.contoso.com/help.html",
        result: {
            resultType: Excel.CustomFunctionValueType.number,
            resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
        },
        parameters: [
            {
                name: "thermometer ID",
                description: "The thermometer to be measured",
                valueType: Excel.CustomFunctionValueType.string,
                valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
            },
            {
                name: "interval (ms)",
                description: "The time between calls",
                valueType: Excel.CustomFunctionValueType.number,
                valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
            },
        ],
        options: { batch: false,  stream: true }
    };
/*
    // sample function that accepts a range of data as its parameter
    function secondHighestTemp(temperatures){ 
        var highest = -273, secondHighest = -273;
        for(var i = 0; i < temperatures.length;i++){
            for(var j = 0; j < temperatures[i].length;j++){
                if(temperatures[i][j] >= highest){
                    secondHighest = highest;
                    highest = temperatures[i][j];
                }
                else if(temperatures[i][j] >= secondHighest){
                    secondHighest = temperatures[i][j];
                }
            }
        }
        return secondHighest;
    }
*/
    Excel.Script.CustomFunctions["CONTOSO"]["SECONDHIGHESTTEMP"] = {
        call: secondHighestTemp,
        description: "Returns the second highest from a range of temperatures",
        helpUrl: "https://example.com/help.html",
        result: {
            resultType: Excel.CustomFunctionValueType.number,
            resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
        },
        parameters: [
            {
                name: "temps",
                description: "the temperatures to be compared",
                valueType: Excel.CustomFunctionValueType.number,
                valueDimensionality: Excel.CustomFunctionDimensionality.matrix,
            },
        ],
        options: { batch: false, stream: false }
    };

    function isPrime(n) {
        var root = Math.sqrt(n);
        if (n < 2) return false;
        for (var divisor = 2; divisor <= root; divisor++){
            if(n % divisor == 0) return false;
        }
        return true;
    }

    Excel.Script.CustomFunctions["CONTOSO"]["ISPRIME"] = {
        call: isPrime,
        description: "Determines whether the input is prime",
        helpUrl: "https://example.com/help.html",
        result: {
            resultType: Excel.CustomFunctionValueType.boolean,
            resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
        },
        parameters: [
            {
                name: "n",
                description: "the number to be evaluated",
                valueType: Excel.CustomFunctionValueType.number,
                valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
            },
        ],
        options: { batch: false, stream: false }
    };
/*
    function getRandom(min, max) {
        return new OfficeExtension.Promise(function(setResult, setError){
            sendRandomOrgHTTP(min, max, function(result){
                if(result.number) setResult(number);
                else setError(result.error);
            });
        });
    }

    Excel.Script.CustomFunctions["CONTOSO"]["RANDOM"] = {
        call: getRandom,
        description: "Generates a random integer between two values",
        helpUrl: "https://example.com/help.html",
        result: {
            resultType: Excel.CustomFunctionValueType.number,
            resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
        },
        parameters: [
            {
                name: "min",
                description: "the number to be evaluated",
                valueType: Excel.CustomFunctionValueType.number,
                valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
            },
            {
                name: "max",
                description: "the number to be evaluated",
                valueType: Excel.CustomFunctionValueType.number,
                valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
            },
        ],
        options: { batch: false, stream: false }
    };
*/
    // register all the functions

    
    Excel.run(function (context) {
        context.workbook.customFunctions.addAll();
        return context.sync().then(function(){});
    }).catch(function(error){});

/*
    // new fn
    Excel.run(function (context) {
        context.workbook.customFunctions.addAll();
        return context.sync().then(function(){

        });
    
    }).catch(function(error){
        console.log("error" + error);
    });
*/
    // Helper functions are below

    // The sendWebRequestExample function simulates a web request to get a temperature
    function sendWebRequestExample(input, callback){
        var result = {};
        // generate a temperature
        result["temperature"] = 42 - (Math.random() * 10);
        setTimeout(function(){
            callback(result);
        }, 250);
    }

    // The log function lets you write debugging messages into Excel (first evaluate the MY.DEBUG function in Excel). You can also debug with regular debugging tools like VS.
    var debug = [];
    var debugUpdate = function(data){};
    function log(myText){
        debug.push([myText]);
        debugUpdate(debug);
    }
    function myDebug(setResult){
        debugUpdate = setResult;
    }
   
}; 