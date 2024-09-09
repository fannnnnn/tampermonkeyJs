// ==UserScript==
// @name         skyward import scores from googleclass, Quizizz
// @namespace    http://tampermonkey.net/
// @version      1.0
// @description  skyward import googleclass points
// @author       YF
// @match        https://skyward.iscorp.com/scripts/wsisa.dll/WService=wseduintleadoftxtx/sepgrb08.w
// @icon         data:image/gif;base64,R0lGODlhAQABAAAAACH5BAEKAAEALAAAAAABAAEAAAICTAEAOw==
// @grant        none
// @updateURL    https://raw.githubusercontent.com/fannnnnn/tampermonkeyJs/main/skywardImportScores.js
// ==/UserScript==


(function () {
    'use strict';

    var script = document.createElement('script');
    script.type = 'text/javascript';
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
    document.head.appendChild(script);

    let controlBarLeft = document.querySelector("#controlBarLeft")
    let div = document.createElement("div");
    div.innerHTML = '<input type="file" id="fileInput" />';
    document.body.append(div);
    controlBarLeft.append(div);

    let notMatch = [];
    let matched = [];

    document.getElementById('fileInput').addEventListener('change', function (e) {
        var file = e.target.files[0];
        var reader = new FileReader();
        // clearScore();
        // excel
        if (file.name.indexOf('xlsx') > 0) {
            var wb;
            var rABS = false;

            reader.onload = function (e) {
                var data = e.target.result;
                if (rABS) {
                    wb = XLSX.read(btoa(fixdata(data)), {
                        type: 'base64'
                    });
                } else {
                    wb = XLSX.read(data, {
                        type: 'binary'
                    });
                }
                let scoreArray = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[1]]);
                for (let i = 0; i < scoreArray.length; i++) {
                    //console.log(' firstName:' + scoreArray[i].m + ' lastName:' + scoreArray[i].x + ' score:' + scoreArray[i].s * 10);
                    setScoreOne(scoreArray[i].m, scoreArray[i].x, scoreArray[i].s * 10, matchNameQuizizz);
                }
                console.log('notMatch');
                console.log(notMatch);
                console.log('matched');
                console.log(matched);
            };
            if (rABS) {
                reader.readAsArrayBuffer(file);
            } else {
                reader.readAsBinaryString(file);
            }
            return;
        }
        // csv
        reader.onload = function (e) {
            var content = e.target.result;
            let textContent = content;
            // console.log(textContent);
            let scoreArray = getLines(textContent);
            setScores(scoreArray);
            console.log('notMatch');
            console.log(notMatch);
            console.log('matched');
            console.log(matched);
        };
        reader.readAsText(file);
    });

    // csv
    function getLines(textContent) {
        textContent = textContent.replace("\r\n", "\n").replace("\r", "\n")
        const scoreArray = textContent.split("\n");
        return scoreArray.slice(1, scoreArray.length);
    }

    function setScores(scoreArray) {
        for (let i = 0; i < scoreArray.length; i++) {
            let lineText = scoreArray[i];
            if (!lineText || lineText.length < 3) {
                continue;
            }
            const infos = lineText.split(",");
            let firstName = infos[1];
            if (firstName) {
                firstName = firstName.trim();
            } else {
                continue;
            }
            let lastName = infos[0];
            if (lastName) {
                lastName = lastName.trim();
            } else {
                continue;
            }
            let score = infos[3];
            setScoreOne(firstName, lastName, score, matchNameGoogleClass);
        }
    }

    // excel
    function fixdata(data) {
        var o = "",
            l = 0,
            w = 10240;
        for (; l < data.byteLength / w; ++l) o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w, l * w + w)));
        o += String.fromCharCode.apply(null, new Uint8Array(data.slice(l * w)));
        return o;
    }

    let uesedIdSet = new Set();
    function setScoreOne(firstName, lastName, score, matchNameF) {
        if (firstName) {
            firstName = firstName.replace('	', ' ').replace('	', ' ').replace('.', '').trim().toUpperCase();
        }
        if (lastName) {
            lastName = lastName.replace('	', ' ').replace('.', '').trim().toUpperCase();
        }
        for (let i = 1; i < 100; i++) {
            let id = "#stu" + i;
            if (uesedIdSet.has(id)) {
                continue;
            }
            let stu = document.querySelector(id);
            if (!stu) {
                break;
            }
            let firstName1 = document.querySelector(id + " > td:nth-child(2)").textContent.trim().replace('  ', ' ')
            let lastName1 = document.querySelector(id + " > td:nth-child(3)").textContent.trim().replace('  ', ' ')
            //console.log(firstName1 + '-' + firstName + '-' + lastName1 + '-' + lastName)
            if (matchNameF(firstName, firstName1, lastName, lastName1)) {
                //console.log('match:' + firstName1 + ' ' + lastName1 + ' ' + score);
                let v = document.querySelector(id + " > td.TD-B.headingLabels.isNbr").childNodes[0];
                if (!v.value || v.value == '' || v.value == '*' || Math.round(score) > v.value) {
                    v.value = Math.round(score);
                }
                let comment = 'match, graded';
                if (!score) {
                    comment = 'match, no score';
                }
                document.querySelector(id + " > td.TD-RB.headingLabels").childNodes[0].value = comment;
                uesedIdSet.add(id);
                matched.unshift(firstName + '-' + firstName1 + '@' + lastName + '-' + lastName1 + '@' + score);
                return;
            }
        }
        notMatch.unshift(firstName + '@' + lastName + '@' + score);
    }

    // Quizizz match name
    function matchNameQuizizz(firstName, firstName1, lastName, lastName1) {
        return (firstName && firstName.length > 1 && (firstName1.indexOf(firstName) >= 0 || firstName.indexOf(firstName1) >= 0)) &&
            (!lastName || lastName == '' || lastName1.indexOf(lastName) >= 0 || lastName.indexOf(lastName1) >= 0
            );
    }

    // GoogleClass
    function matchNameGoogleClass(firstName, firstName1, lastName, lastName1) {
        return (firstName1.indexOf(firstName) === 0 || firstName.indexOf(firstName1) === 0) && lastName1.indexOf(lastName) === 0;
    }

    function clearScore() {
        for (let i = 1; i < 100; i++) {
            let id = "#stu" + i;
            let stu = document.querySelector(id);
            if (!stu) {
                break;
            }
            document.querySelector(id + " > td.TD-B.headingLabels.isNbr").childNodes[0].value = '*';
            document.querySelector(id + " > td.TD-RB.headingLabels").childNodes[0].value = '';
        }
    }

})();