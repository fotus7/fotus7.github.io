const fantasyApp = angular.module("fantasyApp", [])

fantasyApp.controller("mainController", function ($scope, $q, $parse) {
    let companyObjectList = [];
    let productObjectList = [];
    let xslmObjects = [];
  
    $scope.init = function () {
        $scope.activeTab = 'company';
        $scope.activeYear = 1;
        $scope.isDescOrder = true;
        $scope.currentOrderName;
        $scope.quantity = 1000;

        sendRequestByUrl('2022').then(function success(response) {
            if (response) {
                sendRequestByUrl('2021').then(function success(response) {
                    if (response) {
                        sendRequestByUrl('2020').then(function success(response) {
                            $scope.isLoaded = response;
                            $scope.showCompany();
                        });
                    }
                });
            }
        });
    };

    function sendRequestByUrl(year) {
        const deferred = $q.defer();
        const path = `https://fotus7.github.io/lesli/etc/files/otgruzka_${year}.xlsm`;

        let xmlHttpRequest = new XMLHttpRequest();
        xmlHttpRequest.open("GET", path, true);
        xmlHttpRequest.responseType = 'arraybuffer';

        xmlHttpRequest.onload = function () {
            buildXslmObject(xmlHttpRequest, year);
            buildProductObject(year);
            buildCompanyObjectList(year);

            deferred.resolve(true);
        }

        xmlHttpRequest.send();
        return deferred.promise;
    }

    function buildXslmObject(xmlHttpRequest, year) {
        const arrayBuffer = xmlHttpRequest.response;
        const data = new Uint8Array(arrayBuffer);
        const dataArray = new Array();

        for(let i = 0; i != data.length; ++i)  {
            dataArray[i] = String.fromCharCode(data[i]);
        }
        const binaryDataString = dataArray.join("");
        const workbook = XLSX.read(binaryDataString, { type:"binary" });
        const sheetName = workbook.SheetNames[0];

        xslmObjects[year] = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
    }


    function buildProductObject(year) {
        let products = xslmObjects[year].find(object => object['???????? ????????'] === '??????????');
        for (const key of Object.keys(products)) {
            if (key !== '???????? ????????') {
                let product = productObjectList.find(object => object.name === key);
                if (!product) {
                    productObjectList.push({name: key, data: { [year]: { total: 0}}, total: 0});
                }
            }
        }
    }

    function buildCompanyObjectList(year) {
        xslmObjects[year].forEach(element => {
            const otg_date = element['???????? ????????'];
            const month = otg_date ? otg_date.split(' ')[1] : undefined;

            for (const [key, value] of Object.entries(element)) {
                if (element['???????? ????????'] !== '??????????') {
                    const total = parseFloat(value);
                    let product = productObjectList.find(object => object.name === key);
                    if (product) {
                        if (!product.data[year]) {
                            product.data[year] = {};
                            product.data[year].total = 0;
                        }

                        product.data[year][month] = product.data[year][month] ? $scope.round(product.data[year][month], total) : total;
                        product.data[year].total = $scope.round(product.data[year].total, total);
                        product.total = $scope.round(product.total, total);
                    }
                }
            }

            const company = Object.values(element)[1];
            let total = parseFloat(element['??????????']);
            let sameElement = companyObjectList.find(object => object.name === company);

            if (sameElement) {
                if (!sameElement.data[year]) {
                  sameElement.data[year] = {};
                  sameElement.data[year].total = 0;
                }

                sameElement.data[year][month] = sameElement.data[year][month] ? $scope.round(sameElement.data[year][month], total) : total;
                sameElement.data[year].total = $scope.round(sameElement.data[year].total, total);
                sameElement.total = $scope.round(sameElement.total, total);
            }

            if (company && total && month && !sameElement) {
                companyObjectList.push({ name: company, data: { [year]: { [month]: total, total } }, total});
            }
        });
    }

    $scope.setYear = function (year) {
        $scope.isLoaded = false;
        $scope.activeYear = year;
        switch (year) {
            case 1:
                if ($scope.activeTab === 'stats') {
                  $scope.showStats();
                }
                $scope.isLoaded = true;
                break;
            case 3:
                showYearData(true);
                break;
            case 5:
                if (Object.keys(xslmObjects).length < 5) {
                    sendRequestByUrl('2019').then(function success(response) {
                        if (response) {
                            sendRequestByUrl('2018').then(function success(response) {
                                showYearData(response);
                            });
                        }
                    });
                } else {
                    if ($scope.activeTab === 'stats') {
                        $scope.showStats();
                    }
                    $scope.isLoaded = true;
                }
                break;
          default:
        }

        function showYearData(response) {
            if ($scope.activeTab === 'stats') {
                $scope.showStats();
            } else if ($scope.activeTab === 'topCompany') {
                $scope.showTopCompany();
            } else if ($scope.activeTab === 'topProduct') {
                 $scope.showTopProduct();
            } else if ($scope.activeTab === 'analysis') {
                buildAnalysis();
            }
            $scope.isLoaded = response;
        }
    };

    $scope.showCompany = function () {
        clearSearch();
        $scope.currentOrderName = null;
        $scope.activeTab = 'company';
        $scope.quantity = 1000;
        $scope.title = "????????????????";
        $scope.itemList = companyObjectList;
        $scope.copyItemList = companyObjectList;
    };

    $scope.showTopCompany = function () {
        clearSearch();
        $scope.currentOrderName = null;
        $scope.activeTab = 'topCompany';
        $scope.title = "????????????????";
        $scope.quantity = 10;
        $scope.isDescOrder = true;
        $scope.itemList = companyObjectList;
        $scope.orderByField('data[2022].total');

    }

    $scope.showProduct = function () {
        clearSearch();
        $scope.currentOrderName = null;
        $scope.activeTab = 'product';
        $scope.quantity = 1000;
        $scope.title = "??????????????";
        $scope.itemList = productObjectList;
        $scope.copyItemList = productObjectList;
    };

    $scope.showTopProduct = function () {
        clearSearch();
        $scope.currentOrderName = null;
        $scope.activeTab = 'topProduct';
        $scope.title = "??????????????";
        $scope.quantity = 10;
        $scope.isDescOrder = true;
        $scope.itemList = productObjectList;
        $scope.orderByField('data[2022].total');
    }

    $scope.orderByField = function (field) {
        if ($scope.currentOrderName === field) {
            $scope.isDescOrder = !$scope.isDescOrder;
        }

        $scope.currentOrderName = field;
        $scope.itemList.sort(function (a, b) {
          return sort(a, b);
        });

        function sort(a, b) {
            const getter = $parse(field);
            const firstValue = getter(a);
            const secondValue = getter(b);

            if (firstValue === undefined) {
              return 1;
            }
            else if (secondValue === undefined) {
              return -1;
            }
            else if (firstValue === secondValue) {
                return 0;
            }
            else if ($scope.isDescOrder) {
              return firstValue < secondValue ? 1 : -1;
            }
            else {
              return firstValue < secondValue ? -1 : 1;
            }
        }
    };

    $scope.showStats = function () {
        clearSearch();
        $scope.activeTab = 'stats';

        google.charts.load('current', {'packages':['corechart']});
        google.charts.setOnLoadCallback(drawChart);

        function drawChart() {
            let chartData = [
                ['??????????', '??????????, ????']
            ]

            let chartDataArray2022 = [
                ['Jan 2022',  sum(productObjectList, 'Jan', '2022')],
                ['Feb 2022',  sum(productObjectList, 'Feb', '2022')],
                ['Mar 2022',  sum(productObjectList, 'Mar', '2022')],
                ['Apr 2022',  sum(productObjectList, 'Apr', '2022')],
                ['May 2022',  sum(productObjectList, 'May', '2022')],
                ['Jun 2022',  sum(productObjectList, 'Jun', '2022')],
                ['Jul 2022',  sum(productObjectList, 'Jul', '2022')],
                ['Aug 2022',  sum(productObjectList, 'Aug', '2022')],
                ['Sep 2022',  sum(productObjectList, 'Sep', '2022')],
                ['Oct 2022',  sum(productObjectList, 'Oct', '2022')],
                ['Nov 2022',  sum(productObjectList, 'Nov', '2022')],
                ['Dec 2022',  sum(productObjectList, 'Dec', '2022')]
            ];

            let chartDataArray20202021 = [
                ['Jan 2020', sum(productObjectList, 'Jan', '2020')],
                ['Feb 2020', sum(productObjectList, 'Feb', '2020')],
                ['Mar 2020',  sum(productObjectList, 'Mar', '2020')],
                ['Apr 2020',  sum(productObjectList, 'Apr', '2020')],
                ['May 2020',  sum(productObjectList, 'May', '2020')],
                ['Jun 2020',  sum(productObjectList, 'Jun', '2020')],
                ['Jul 2020',  sum(productObjectList, 'Jul', '2020')],
                ['Aug 2020',  sum(productObjectList, 'Aug', '2020')],
                ['Sep 2020',  sum(productObjectList, 'Sep', '2020')],
                ['Oct 2020',  sum(productObjectList, 'Oct', '2020')],
                ['Nov 2020',  sum(productObjectList, 'Nov', '2020')],
                ['Dec 2020',  sum(productObjectList, 'Dec', '2020')],
                ['Jan 2021',  sum(productObjectList, 'Jan', '2021')],
                ['Feb 2021',  sum(productObjectList, 'Feb', '2021')],
                ['Mar 2021',  sum(productObjectList, 'Mar', '2021')],
                ['Apr 2021',  sum(productObjectList, 'Apr', '2021')],
                ['May 2021',  sum(productObjectList, 'May', '2021')],
                ['Jun 2021',  sum(productObjectList, 'Jun', '2021')],
                ['Jul 2021',  sum(productObjectList, 'Jul', '2021')],
                ['Aug 2021',  sum(productObjectList, 'Aug', '2021')],
                ['Sep 2021',  sum(productObjectList, 'Sep', '2021')],
                ['Oct 2021',  sum(productObjectList, 'Oct', '2021')],
                ['Nov 2021',  sum(productObjectList, 'Nov', '2021')],
                ['Dec 2021',  sum(productObjectList, 'Dec', '2021')]
            ];

            let chartDataArray20182019 = [
                ['Jan 2018', sum(productObjectList, 'Jan', '2018')],
                ['Feb 2018', sum(productObjectList, 'Feb', '2018')],
                ['Mar 2018',  sum(productObjectList, 'Mar', '2018')],
                ['Apr 2018',  sum(productObjectList, 'Apr', '2018')],
                ['May 2018',  sum(productObjectList, 'May', '2018')],
                ['Jun 2018',  sum(productObjectList, 'Jun', '2018')],
                ['Jul 2018',  sum(productObjectList, 'Jul', '2018')],
                ['Aug 2018',  sum(productObjectList, 'Aug', '2018')],
                ['Sep 2018',  sum(productObjectList, 'Sep', '2018')],
                ['Oct 2018',  sum(productObjectList, 'Oct', '2018')],
                ['Nov 2018',  sum(productObjectList, 'Nov', '2018')],
                ['Dec 2018',  sum(productObjectList, 'Dec', '2018')],
                ['Jan 2019',  sum(productObjectList, 'Jan', '2019')],
                ['Feb 2019',  sum(productObjectList, 'Feb', '2019')],
                ['Mar 2019',  sum(productObjectList, 'Mar', '2019')],
                ['Apr 2019',  sum(productObjectList, 'Apr', '2019')],
                ['May 2019',  sum(productObjectList, 'May', '2019')],
                ['Jun 2019',  sum(productObjectList, 'Jun', '2019')],
                ['Jul 2019',  sum(productObjectList, 'Jul', '2019')],
                ['Aug 2019',  sum(productObjectList, 'Aug', '2019')],
                ['Sep 2019',  sum(productObjectList, 'Sep', '2019')],
                ['Oct 2019',  sum(productObjectList, 'Oct', '2019')],
                ['Nov 2019',  sum(productObjectList, 'Nov', '2019')],
                ['Dec 2019',  sum(productObjectList, 'Dec', '2019')]
            ];

            let year = "2022";

            if ($scope.activeYear === 3) {
                chartDataArray2022 = chartDataArray20202021.concat(chartDataArray2022);
                year = "2020 - 2022";
            }

            if ($scope.activeYear === 5) {
                chartDataArray2022 = chartDataArray20202021.concat(chartDataArray2022);
                chartDataArray2022 = chartDataArray20182019.concat(chartDataArray2022);
                year = "2018 - 2022";
            }

            chartData = chartData.concat(chartDataArray2022);
            const data = google.visualization.arrayToDataTable(chartData);

            const options = {
                title: `???????????????????? ????????????, ${year}`,
                curveType: 'line',
                legend: { position: 'bottom' },
            };

            const chart = new google.visualization.LineChart(document.getElementById('curve_chart'));
            chart.draw(data, options);
        }
    };

    $scope.round = function (firstValue, secondValue) {
        if (!secondValue) secondValue = 0;
        return Math.round((firstValue + secondValue) * 10) / 10;
    }

    $scope.showAnalysis = function () {
        clearSearch();
        $scope.activeTab = 'analysis';
        $scope.setYear(3);
    }

    function buildAnalysis() {
        const periodList = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'];
        let analysisList = [];
        analysisList = buildAnalysisList(2022, analysisList);
        analysisList = buildAnalysisList(2021, analysisList);

        let resultList = [];
        analysisList.forEach(element => {
            let currentYearPeriodValue = 0;
            let previousYearPeriodValue = 0;
            if (element[2021]) {
                periodList.forEach((month) => {
                    currentYearPeriodValue += element[2022]?.[month] ? element[2022][month] : 0;
                    previousYearPeriodValue += element[2021][month] ? element[2021][month] : 0;
                });
            }
            element.percent = $scope.round((currentYearPeriodValue - previousYearPeriodValue) * 100 / previousYearPeriodValue);
            if (element.percent <= -25) {
                resultList.push({ productName: element.productName, companyName: element.companyName, percent: element.percent, '2021': Math.round(previousYearPeriodValue * 10) / 10, '2022': Math.round(currentYearPeriodValue * 10) / 10 });
            }
        });

        $scope.analysisList = resultList;
    }

    function buildAnalysisList(year, analysisList) {
        xslmObjects[year].forEach(element => {
            const companyName = element['????????????????'];
            Object.keys(element).forEach(key => {
                if (key !== '???????? ????????' && key !== '????????????????' && key !== '??????????' && key !== '?????? ????????????') {
                    const productName = key;
                    const total = parseFloat(element[productName]);
                    const otg_date = element['???????? ????????'];
                    const month = otg_date ? otg_date.split(' ')[1] : undefined;

                    let existedObject = analysisList.find(object => object.companyName === companyName && object.productName === productName);
                    if (existedObject) {
                        if (existedObject[year]) {
                            existedObject[year][month] = $scope.round(total, existedObject[year][month]);
                        } else {
                            existedObject[year] = { [month]: total };
                        }
                    }

                    if (companyName && key !== '??????????' && total && !existedObject) {
                        analysisList.push({companyName, productName, [year]: { [month]: total} });
                    }
                }
            });
        });

        return analysisList;
    }

    $scope.search = function (searchValue) {
        if (searchValue) {
            let filteredItemList = [];

            $scope.copyItemList.forEach(item => {
                if (item.name.toLowerCase().includes(searchValue.toLowerCase())) {
                    filteredItemList.push(item);
                    return;
                }
            });

            $scope.itemList = filteredItemList;
        } else {
            $scope.itemList = $scope.copyItemList;
        }
    };

    $scope.showDetails = function (detailName, month, year) {
        $scope.details = {name: detailName, month, year, items: []};
        if (year) {
            xslmObjects[year].forEach(element => {
                buildDetailsJson(element);
            });
        } else {
            xslmObjects.forEach((yearElement, index) => {
                if ($scope.activeYear === 5 || ($scope.activeYear === 3 && index >= 2020))
                yearElement.forEach(element => {
                    buildDetailsJson(element);
                });
            });
        }

        $('#detailsPopup').show();

        function buildDetailsJson(element) {
            switch($scope.activeTab) {
                case 'topProduct':
                case 'product':
                    if (element[detailName] && (month === '????????' || element['???????? ????????'].includes(month))) {
                        let item = $scope.details.items.find((item) => item.name === element['????????????????']);
                        if (item) {
                            item.value = $scope.round(parseFloat(item.value), parseFloat(element[detailName]))
                        } else if (element['????????????????']) {
                            $scope.details.items.push({ name: element['????????????????'], value: element[detailName] });
                        }
                    }

                    break;
                case 'topCompany':
                case 'company':
                    if (element['????????????????'] === detailName && (month === '????????' || element['???????? ????????'].includes(month))) {
                        Object.keys(element).forEach(key => {
                            if (key !== '???????? ????????' && key !== '??????????' && key !== '????????????????') {
                                let item = $scope.details.items.find((item) => item.name === key);
                                if (item) {
                                    item.value = $scope.round(parseFloat(item.value), parseFloat(element[key]))
                                } else {
                                    $scope.details.items.push({ name: key, value: element[key] });
                                }
                            }
                        });
                    }

                    break;
            }
        }
    };

    $scope.closeDetailsPopup = function () {
        $('#detailsPopup').hide();
    }

    function clearSearch() {
        $scope.searchValue = '';
        $scope.itemList = $scope.copyItemList;
    }

    function sum(items, prop, year){
        return items.reduce(function(a, b) {
            const t = b.data[year] && b.data[year][prop] ? b.data[year][prop] : 0;
            return $scope.round(a, t);
        }, 0);
    };
});