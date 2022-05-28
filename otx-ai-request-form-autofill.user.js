// ==UserScript==
// @name         otx-ai-request-form-autofill
// @version      0.3
// @description  Make AI Creation a little bit easier :)
// @author       Lindjunne Gerard Montenegro II (lmontene@opentext.com)
// @match        https://forms.office.com/pages/responsepage.aspx?id=d4ShEDPVzU6njZFtvYSdfBJ9v22uKC5Lt-HYhNFCpwlUNURWSEwxQlNDVklJWEs4TlpJRVdNWE1NQi4u*
// @grant        none
// @downloadURL  https://raw.githubusercontent.com/SiNONiMiTY/otx-ai-request-form-autofill/main/otx-ai-request-form-autofill.user.js
// ==/UserScript==



// DO NOT MODIFY BELOW THIS LINE
// =============================
(function() {
    "use strict";

    const USERSCRIPT_UUID = "{dfd847b5-b90d-45e7-b039-467dd09ab2b7}";

    /**
     * Retrieve user preferences from local storage if present.
     * If not, prompt user to provide values.
     */
    let formGeoLead = localStorage.getItem( `${USERSCRIPT_UUID}_geoLead` ),
        formTeamName = localStorage.getItem( `${USERSCRIPT_UUID}_teamName` ),
        formRequestMode = localStorage.getItem( `${USERSCRIPT_UUID}_requestMode` );

    while ( !formGeoLead ) {
        formGeoLead = prompt( "Geo Lead (Director)", "" );
        localStorage.setItem( `${USERSCRIPT_UUID}_geoLead`, formGeoLead );
    }

    while ( !formTeamName ) {
        formTeamName = prompt( "Team Name", "" );
        localStorage.setItem( `${USERSCRIPT_UUID}_teamName`, formTeamName );
    }

    while ( !formRequestMode ) {
        formRequestMode = prompt( "Request Mode [single, bulk]", "" );
        localStorage.setItem( `${USERSCRIPT_UUID}_requestMode`, formRequestMode );
    }

    /**
     * Evaluate user preferences and automate the form fill-up procedure.
     */
    let previousQuestionText;

    if ( formGeoLead && formTeamName && formRequestMode ) {
        document.addEventListener( "DOMNodeInserted", function() {
            let questionTextParentNode = document.getElementsByClassName( "office-form-question-title" ),
                questionTextChildNode = questionTextParentNode[0].getElementsByClassName( "text-format-content" ),
                questionText = questionTextChildNode[0].innerHTML;

            if ( questionText && questionText !== previousQuestionText ) {
                previousQuestionText = questionText;

                switch ( questionText.toUpperCase() ) {
                    case "SELECT GEO":
                        automateResponse( "r2c7dbec1c3fa4ef5a1c1ba9fc4722f47", formGeoLead, true );
                        break;

                    case "SELECT EAJ TEAM":
                        automateResponse( "r6632b90930654b89809ba3bb2b6620d2", formTeamName, true );
                        break;

                    case "SELECT US NR TEAM":
                        automateResponse( "r81ea9fb4ad5e4b6db3b9e313fe53a277", formTeamName, true );
                        break;

                    case "SELECT ART TEAM":
                        automateResponse( "r54d92d15fdf7464a903283b2852e05d1", formTeamName, true );
                        break;

                    case "SELECT US VBAG TEAM":
                        automateResponse( "r4d53bc83b8914a2c8279514b82a1983a", formTeamName, true );
                        break;

                    case "SELECT FIRE TEAM":
                        automateResponse( "r3a8c626551fe433ba9aa8a32edc47a61", formTeamName, true );
                        break;

                    case "SELECT REQUEST MODE":
                        automateResponse( "rfe94edb4312b47bbb280c02ad5b650b8", formRequestMode, true );
                        break;
                }
            }
        }, false );
    }

    function automateResponse( questionTextKey, answer, proceedToNextPage = false ) {
        let questionTextValue = document.getElementsByName( questionTextKey ),
            btn_Next = document.getElementsByClassName( "section-next-button" );

        for ( let i = 0; i < questionTextValue.length; i++ ) {
            if ( questionTextValue[i].value.toUpperCase() === answer.toUpperCase() ) {
                questionTextValue[i].click();
                proceedToNextPage ? btn_Next[0].click() : false;
            }
        }

        return;
    }
})();
