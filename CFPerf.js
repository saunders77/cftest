var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t;
    return { next: verb(0), "throw": verb(1), "return": verb(2) };
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = y[op[0] & 2 ? "return" : op[0] ? "throw" : "next"]) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [0, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Excel.Script.CustomFunctions = {
    "msft.perf": {
        "mortgagePaymentJS": {
            call: mortgagePaymentJS,
            description: "Computes the mortgage payment",
            helpUrl: "http://dev.office.com",
            result: {
                resultType: Excel.CustomFunctionValueType.number,
                resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
            },
            parameters: [
                {
                    name: "principalAmount",
                    description: "Number",
                    valueType: Excel.CustomFunctionValueType.number,
                    valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
                },
                {
                    name: "interestRate",
                    description: "Number",
                    valueType: Excel.CustomFunctionValueType.number,
                    valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
                },
                {
                    name: "numberOfMonths",
                    description: "Number",
                    valueType: Excel.CustomFunctionValueType.number,
                    valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
                },
            ],
            options: {
                batch: false,
                stream: false,
                cancelable: false,
            }
        },
        "findNthPrimeJS": {
            call: findNthPrimeJS,
            description: "Finds the nth prime",
            helpUrl: "http://dev.office.com",
            result: {
                resultType: Excel.CustomFunctionValueType.number,
                resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
            },
            parameters: [
                {
                    name: "n",
                    description: "Nth prime",
                    valueType: Excel.CustomFunctionValueType.number,
                    valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
                },
            ],
            options: {
                batch: false,
                stream: false,
                cancelable: false,
            }
        },
        "bubbleSortJS": {
            call: bubbleSortJS,
            description: "Finds how many swaps will be performed to sort the input numbers",
            helpUrl: "http://dev.office.com",
            result: {
                resultType: Excel.CustomFunctionValueType.number,
                resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
            },
            parameters: [
                {
                    name: "numbers",
                    description: "Numbers to perform bubble sort on",
                    valueType: Excel.CustomFunctionValueType.number,
                    valueDimensionality: Excel.CustomFunctionDimensionality.matrix,
                },
            ],
            options: {
                batch: false,
                stream: false,
                cancelable: false,
            }
        },
        "slowAddBillionJS": {
            call: slowAddBillionJS,
            description: "Add billion in the slowest way",
            helpUrl: "http://dev.office.com",
            result: {
                resultType: Excel.CustomFunctionValueType.number,
                resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
            },
            parameters: [
                {
                    name: "n",
                    description: "Add billiion to n",
                    valueType: Excel.CustomFunctionValueType.number,
                    valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
                },
            ],
            options: {
                batch: false,
                stream: false,
                cancelable: false,
            }
        },
        "AddArrayNumbersAndNumber": {
            call: AddArrayNumbersAndNumber,
            description: "Adds numbers in a range and a number in a cell",
            helpUrl: "http://dev.office.com",
            result: {
                resultType: Excel.CustomFunctionValueType.number,
                resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
            },
            parameters: [
                {
                    name: "array",
                    description: "First",
                    valueType: Excel.CustomFunctionValueType.number,
                    valueDimensionality: Excel.CustomFunctionDimensionality.matrix,
                },
                {
                    name: "number",
                    description: "Second",
                    valueType: Excel.CustomFunctionValueType.number,
                    valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
                },
            ],
            options: {
                batch: false,
                stream: false,
                cancelable: false,
            }
        },
        "ReturnArrayNumbers": {
            call: ReturnArrayNumbers,
            description: "Take numbers in a range and return those numbers",
            helpUrl: "http://dev.office.com",
            result: {
                resultType: Excel.CustomFunctionValueType.number,
                resultDimensionality: Excel.CustomFunctionDimensionality.matrix,
            },
            parameters: [
                {
                    name: "array",
                    description: "First",
                    valueType: Excel.CustomFunctionValueType.number,
                    valueDimensionality: Excel.CustomFunctionDimensionality.matrix,
                },
            ],
            options: {
                batch: false,
                stream: false,
                cancelable: false,
            }
        },
    },
};
Office.initialize = function (reason) {
    window.Promise = OfficeExtension.Promise;
    Excel.run(function (context) {
        context.workbook.customFunctions.addAll();
        return context.sync();
    });
};
function swap(numbers, i, j) {
    var temp = numbers[i];
    numbers[i] = numbers[j];
    numbers[j] = temp;
}
function mortgagePaymentJS(principalAmount, interestRate, numberOfMonths) {
    interestRate = (interestRate / 100) / 12;
    var irr = Math.pow(1 + interestRate, numberOfMonths);
    return interestRate * principalAmount * irr / (irr - 1);
}
function bubbleSortJS(array) {
    var numbers = [];
    for (var i = 0; i < array.length; ++i) {
        for (var j = 0; j < array[i].length; ++j) {
            numbers.push(array[i][j]);
        }
    }
    var numberOfSwaps = 0;
    var length = numbers.length;
    for (var i = 0; i < length - 1; i++) {
        for (var j = 0; j < length - i - 1; j++) {
            if (numbers[j] > numbers[j + 1]) {
                numberOfSwaps++;
                swap(numbers, j, j + 1);
            }
        }
    }
    return numberOfSwaps;
}

function isPrimeNumber(current, vectorPrimes) {
    var length = vectorPrimes.length;
    for (var i = 0; i < length; i++) {
        if ((current % vectorPrimes[i]) == 0) {
            return false;
        }
    }
    return true;
}

function findNthPrimeJS(n) {
   if (n <= 0) {
        return 0;
    }
    var count = 0;
    var current = 1;
    var vectorPrimes = [];
    do {
        current++;
        if (isPrimeNumber(current, vectorPrimes)) {
            vectorPrimes.push(current);
            count++;
        }
    } while (count < n);
    return current;
}

function slowAddBillionJS(num) {
  var numberOfOnes = 1000000000;
  while (numberOfOnes != 0) {
    num = num + 1;
    numberOfOnes--;
  }
  return num;
}

function ReturnArrayNumbers(unused) {
    var items = [
        [1, 2],
        [3, 4],
        [5, 6]
    ];
    return items;
}

function AddArrayNumbersAndNumber(array, num) {
    var sum = 0;
    var i = 0;
    var j = 0;
    for (i = 0; i < array.length; ++i) {
        for (j = 0; j < array[i].length; ++j) {
            sum += array[i][j];
        }
    }
    return sum;
}


function customFunctions_createTestData() {
    return __awaiter(this, void 0, void 0, function () {
        var _this = this;
        return __generator(this, function (_a) {
            Excel.run(function (ctx) { return __awaiter(_this, void 0, void 0, function () {
                var sheetName, sheets, sheet, entireSheet, countOfNumbers, numbers, i, range;
                return __generator(this, function (_a) {
                    switch (_a.label) {
                        case 0:
                            logger.clear();
                            logger.comment("Creating test data ...");
                            sheetName = "PerfData";
                            sheets = ctx.workbook.worksheets;
                            sheet = ctx.workbook.worksheets.getItemOrNullObject(sheetName);
                            ctx.load(sheet);
                            return [4 /*yield*/, ctx.sync()];
                        case 1:
                            _a.sent();
                            if (!sheet.isNull) return [3 /*break*/, 3];
                            sheet = sheets.add(sheetName);
                            ctx.load(sheet);
                            return [4 /*yield*/, ctx.sync()];
                        case 2:
                            _a.sent();
                            _a.label = 3;
                        case 3:
                            entireSheet = sheet.getRange(null);
                            entireSheet.clear(null);
                            return [4 /*yield*/, ctx.sync];
                        case 4:
                            _a.sent();
                            countOfNumbers = 10000;
                            numbers = [];
                            for (i = countOfNumbers; i > 0; i--) {
                                numbers.push([i]);
                            }
                            range = sheet.getRange('A1:A' + countOfNumbers);
                            range.values = numbers;
                            return [4 /*yield*/, ctx.sync()];
                        case 5:
                            _a.sent();
                            logger.comment("Test data has been created.");
                            return [2 /*return*/];
                    }
                });
            }); });
            return [2 /*return*/];
        });
    });
}
var logger = {};
window.onload = function () {
    var loggerElement = document.getElementById('log');
    logger.comment = function () {
        for (var i = 0; i < arguments.length; i++) {
            if (typeof arguments[i] == 'object') {
                loggerElement.innerHTML += (JSON && JSON.stringify ? JSON.stringify(arguments[i], undefined, 2) : arguments[i]) + '<br />';
            }
            else {
                loggerElement.innerHTML += arguments[i] + '<br />';
            }
        }
    };
    logger.clear = function () {
        loggerElement.innerHTML = '';
    };
};
function runTest() {
    var _this = this;
    Excel.run(function (ctx) { return __awaiter(_this, void 0, void 0, function () {
        var numberOfIterations, numberOfFunctions, timeout, selectedFunction, radios, i, radioButton, functionName, functionParamters, bubbleSortNumbers, nthPrime, principalAmount, interestRate, numberOfMonths, err_1;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    _a.trys.push([0, 2, , 3]);
                    logger.clear();
                    numberOfIterations = parseInt(document.getElementById('NumberOfIterations').value);
                    numberOfFunctions = parseInt(document.getElementById('NumberOfFunctions').value);
                    timeout = parseInt(document.getElementById('Timeout').value);
                    selectedFunction = "0";
                    radios = document.getElementsByName('FunctionToRun');
                    for (i = 0; i < radios.length; i++) {
                        radioButton = radios[i];
                        if (radioButton.checked) {
                            selectedFunction = radioButton.value;
                            break;
                        }
                    }
                    functionName = '';
                    functionParamters = '';
                    switch (selectedFunction) {
                        case "1":
                            functionName = "msft.perf.bubbleSortJS";
                            bubbleSortNumbers = document.getElementById('BubbleSortNumbers').value;
                            functionParamters = "(PerfData!A1:A" + bubbleSortNumbers + ")";
                            break;
                        case "2":
                            functionName = "msft.perf.findNthPrimeJS";
                            nthPrime = document.getElementById('NthPrime').value;
                            functionParamters = "(" + nthPrime + ")";
                            break;
                        default:
                            functionName = "msft.perf.mortgagePaymentJS";
                            principalAmount = document.getElementById('PrincipalAmount').value;
                            interestRate = document.getElementById('InterestRate').value;
                            numberOfMonths = document.getElementById('NumberOfMonths').value;
                            functionParamters = "(" + principalAmount + "," + interestRate + "," + numberOfMonths + ")";
                            break;
                    }
                    logger.comment("Calling =" + functionName + functionParamters);
                    return [4 /*yield*/, customFunctions_PerfHelper(ctx, numberOfIterations, numberOfFunctions, functionName, functionParamters, timeout)];
                case 1:
                    _a.sent();
                    return [3 /*break*/, 3];
                case 2:
                    err_1 = _a.sent();
                    logger.comment("Failure occurred: " + err_1);
                    return [3 /*break*/, 3];
                case 3: return [2 /*return*/];
            }
        });
    }); });
    return 0;
}
;
function promisify(action) {
    return new OfficeExtension.Promise(function (resolve, reject) {
        var callback = function (result) {
            if (result.status === "succeeded") {
                resolve(result.value);
            }
            else {
                reject(result.error);
            }
        };
        action(callback);
    });
}
function customFunctions_Execute(ctx, arrayOfFormulas, expectedTimeout) {
    return __awaiter(this, void 0, void 0, function () {
        var _this = this;
        var testApi, endTimestamp, beginTimestamp, beginEventFired, endEventFired, numberOfFunctions, entireSheet;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    testApi = ctx.workbook.internalTest;
                    beginEventFired = false;
                    endEventFired = false;
                    numberOfFunctions = 1000; //arrayOfFormulas.length;
                    // unregister all custom function events
                    testApi.unregisterAllCustomFunctionExecutionEvents();
                    entireSheet = ctx.workbook.worksheets.getItem('Sheet1').getRange(null);
                    entireSheet.clear(null);
                    return [4 /*yield*/, ctx.sync()];
                case 1:
                    _a.sent();
                    return [2 /*return*/, promisify(function (callback) { return __awaiter(_this, void 0, void 0, function () {
                            var _this = this;
                            var beginEvent, timeoutHandle, endEvent, range;
                            return __generator(this, function (_a) {
                                switch (_a.label) {
                                    case 0:
                                        beginEvent = testApi.onCustomFunctionExecutionBeginEvent.add(function (eventArgs) {
                                            beginTimestamp = customFunctions_ConvertTicksToMicroseconds(eventArgs.higherTicks, eventArgs.lowerTicks);
                                            beginEvent.remove();
                                            beginEventFired = true;
                                            return ctx.sync();
                                        });
                                        endEvent = testApi.onCustomFunctionExecutionEndEvent.add(function (eventArgs) {
                                            clearTimeout(timeoutHandle);
                                            endTimestamp = customFunctions_ConvertTicksToMicroseconds(eventArgs.higherTicks, eventArgs.lowerTicks);
                                            endEvent.remove();
                                            endEventFired = true;
                                            var executionTimeMicroseconds = 0;
                                            if (beginEventFired) {
                                                executionTimeMicroseconds = endTimestamp - beginTimestamp;
                                            }
                                            callback({ status: 'succeeded', value: executionTimeMicroseconds });
                                            return ctx.sync();
                                        });
                                        return [4 /*yield*/, ctx.sync()];
                                    case 1:
                                        _a.sent();
                                        range = ctx.workbook.worksheets.getItem('Sheet1').getRange('A1:T50');
                                        range.formulas = arrayOfFormulas;
                                        timeoutHandle = setTimeout(function () { return __awaiter(_this, void 0, void 0, function () {
                                            return __generator(this, function (_a) {
                                                switch (_a.label) {
                                                    case 0:
                                                        if (!beginEventFired) {
                                                            beginEvent.remove();
                                                        }
                                                        endEvent.remove();
                                                        return [4 /*yield*/, ctx.sync()];
                                                    case 1:
                                                        _a.sent();
                                                        callback({ status: 'failed', error: "Customfunction execution event(s) were not fired, CustomFunctionExecutionBeginEvent=" + beginEventFired + ", CustomFunctionExecutionEndEvent=" + endEventFired });
                                                        return [2 /*return*/];
                                                }
                                            });
                                        }); }, expectedTimeout);
                                        return [4 /*yield*/, ctx.sync()];
                                    case 2:
                                        _a.sent();
                                        return [2 /*return*/];
                                }
                            });
                        }); })];
            }
        });
    });
}
function customFunctions_PerfHelper(ctx, numberOfIterations, numberOfFunctions, functionName, functionParameters, expectedTimeout) {
    return __awaiter(this, void 0, void 0, function () {
        var arrayOfFormulas, row, col, rowValue, i_1, executionTimes, i, executionTime, medianExecutionTime;
        return __generator(this, function (_a) {
            switch (_a.label) {
                case 0:
                    arrayOfFormulas = [];
                    //for (i_1 = 0; i_1 < numberOfFunctions; i_1++) {
                    //    arrayOfFormulas.push(['=' + functionName + functionParameters]);
                    //}
                    for (row = 0; row < 50; row++) {
                      rowValue = [];
                      for (col = 0; col < 20; col++) {
                        //var num = row * 20 + col
                        rowValue.push('=' + functionName + functionParameters);
                      }
                      arrayOfFormulas.push(rowValue);
                    }

                    executionTimes = [];
                    i = 0;
                    _a.label = 1;
                case 1:
                    if (!(i < numberOfIterations)) return [3 /*break*/, 4];
                    return [4 /*yield*/, customFunctions_Execute(ctx, arrayOfFormulas, expectedTimeout)];
                case 2:
                    executionTime = (_a.sent()) / 1000;
                    logger.comment("Iteration " + (i + 1) + ": Total time" + executionTime + " ms");
                    executionTimes.push(executionTime);
                    _a.label = 3;
                case 3:
                    i++;
                    return [3 /*break*/, 1];
                case 4:
                    executionTimes.sort();
                    medianExecutionTime = 0;
                    if (numberOfIterations % 2 == 0) {
                        medianExecutionTime = (executionTimes[(numberOfIterations / 2) - 1] + executionTimes[numberOfIterations / 2]) / 2;
                    }
                    else {
                        medianExecutionTime = executionTimes[(numberOfIterations - 1) / 2];
                    }
                    logger.comment("Median = " + medianExecutionTime + " ms");
                    return [2 /*return*/, medianExecutionTime];
            }
        });
    });
}
function customFunctions_ConvertTicksToMicroseconds(higherTocks, lowerTicks) {
    // We cannot use << here because, the result of << is always a 32bit integer
    var microseconds = Math.pow(2, 31) * higherTocks + lowerTicks;
    return microseconds;
}
