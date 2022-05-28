// ==UserScript==
// @name         otx-ai-request-form-autofill
// @version      0.4
// @description  Make AI Creation a little bit easier :)
// @author       Lindjunne Gerard Montenegro II (lmontene@opentext.com)
// @match        https://forms.office.com/pages/responsepage.aspx?id=d4ShEDPVzU6njZFtvYSdfBJ9v22uKC5Lt-HYhNFCpwlUNURWSEwxQlNDVklJWEs4TlpJRVdNWE1NQi4u*
// @grant        none
// @downloadURL  https://raw.githubusercontent.com/SiNONiMiTY/otx-ai-request-form-autofill/release/otx-ai-request-form-autofill.user.js
// ==/UserScript==



(function() {
    "use strict";

    const ACTIVE_SECTION = "officeforms.activesectiond4ShEDPVzU6njZFtvYSdfBJ9v22uKC5Lt-HYhNFCpwlUNURWSEwxQlNDVklJWEs4TlpJRVdNWE1NQi4u.3e61bda7-1985-4adb-9f20-ae92c2e75bfa";
    const ANSWER_MAP = "officeforms.answermap.d4ShEDPVzU6njZFtvYSdfBJ9v22uKC5Lt-HYhNFCpwlUNURWSEwxQlNDVklJWEs4TlpJRVdNWE1NQi4u.3e61bda7-1985-4adb-9f20-ae92c2e75bfa";
    const ANSWER_MAP_TEMPLATE = "dfd847b5-b90d-45e7-b039-467dd09ab2b7_ANSWER_MAP";

    let answerMap = [
        {
            "QuestionId": "r2c7dbec1c3fa4ef5a1c1ba9fc4722f47",
            "QuestionText" : "SELECT GEO"
        },
        {
            "QuestionId": "r81ea9fb4ad5e4b6db3b9e313fe53a277",
            "QuestionText" : "SELECT US NR TEAM"
        },
        {
            "QuestionId": "r3a8c626551fe433ba9aa8a32edc47a61",
            "QuestionText" : "SELECT FIRE TEAM"
        },
        {
            "QuestionId": "r6632b90930654b89809ba3bb2b6620d2",
            "QuestionText" : "SELECT EAJ TEAM"
        },
        {
            "QuestionId": "r4d53bc83b8914a2c8279514b82a1983a",
            "QuestionText" : "SELECT US VBAG TEAM"
        },
        {
            "QuestionId": "r54d92d15fdf7464a903283b2852e05d1",
            "QuestionText" : "SELECT ART TEAM"
        },
        {
            "QuestionId": "rfe94edb4312b47bbb280c02ad5b650b8",
            "QuestionText" : "SELECT REQUEST MODE"
        },
        {
            "QuestionId": "ra2db48cc749e440d8ca956d76996f4f6",
            "QuestionText" : "REQUEST OPTION"
        },
        {
            "QuestionId": "rbbada99ab3b64b8ba3defed5f7e7bae5",
            "QuestionText" : "LOG NUMBER"
        },
        {
            "QuestionId": "r8ac4eb4b712b427590b022164df20134",
            "QuestionText" : "CLIENT CODE"
        },
        {
            "QuestionId": "ra2e3866ee59c4476aebf4b2801291d36",
            "QuestionText" : "DESCRIPTION"
        },
        {
            "QuestionId": "r35ebc531b10949f3a36331b23a75ca7f",
            "QuestionText" : "WOTR"
        },
        {
            "QuestionId": "r992f06d89476465f8250555c2eb62355",
            "QuestionText" : "MAP NAME"
        },
        {
            "QuestionId": "r16a4f2cbf38b4bd9addfe378f8b2b35b",
            "QuestionText" : "ASSIGN TO"
        },
        {
            "QuestionId": "r82aa4e41f87e45ec8aba419139514cbf",
            "QuestionText" : "REQUEST (USER FIELD 1)"
        },
        {
            "QuestionId": "rcadee83420ad49fda7b77261ab777e05",
            "QuestionText" : "RESOLUTION (USER FIELD 2)"
        },
        {
            "QuestionId": "r274594b6550648d4b7ee8ef350e50acf",
            "QuestionText" : "ADD SUPPLEMENTARY DETAILS?"
        },
        {
            "QuestionId": "r94fc321265d24db1892db399032349c9",
            "QuestionText" : ""
        },
        {
            "QuestionId": "re7df054769ee41f5b7284263b4ebeaa5",
            "QuestionText" : ""
        },
        {
            "QuestionId": "r6328a374eeed4730ba26eae394ba67cb",
            "QuestionText" : ""
        },
        {
            "QuestionId": "r14b2180f117b4f09ac473ad3bc4a1c2a",
            "QuestionText" : ""
        },
        {
            "QuestionId": "r10b04ce502e6461983305d0e6549c236",
            "QuestionText" : ""
        },
        {
            "QuestionId": "r9bda24b72cc8484db7a16492630c23bb",
            "QuestionText" : ""
        },
        {
            "QuestionId": "r81c514f938664d2bbee14303d2f86adf",
            "QuestionText" : ""
        },
        {
            "QuestionId": "r8f2cb17dd30c44f8a1929c53c39a8b01",
            "QuestionText" : ""
        },
        {
            "QuestionId": "r374bd5e590f04c269eab47794ebbf185",
            "QuestionText" : ""
        },
        {
            "QuestionId": "r5a629dc22c284fcb9712d122bea96f6c",
            "QuestionText" : ""
        },
        {
            "QuestionId": "r0fef383dddaf4abf8989c7da38c75f10",
            "QuestionText" : ""
        },
        {
            "QuestionId": "r955c951e0bf04df3b36ce8c4e35dcdec",
            "QuestionText" : ""
        },
        {
            "QuestionId": "rde31db4b5221462b8a6231395fe1eaaf",
            "QuestionText" : ""
        },
        {
            "QuestionId": "rfc93a9d1716f4e3c8607383b9da47271",
            "QuestionText" : ""
        },
        {
            "QuestionId": "r91a29801575d4b44ae4df14f4dceea63",
            "QuestionText" : ""
        },
        {
            "QuestionId": "r909754d40b1743cb8d947bd3fd5f2a67",
            "QuestionText" : ""
        },
        {
            "QuestionId": "r6d50f66e3ab645cbb3fe7f129f701076",
            "QuestionText" : ""
        },
        {
            "QuestionId": "r62f976140662462ca5626cd13a553fb3",
            "QuestionText" : ""
        }
    ];

    let questionText;

    // Clean-up current state
    localStorage.removeItem( ACTIVE_SECTION );
    localStorage.removeItem( ANSWER_MAP );

    if ( localStorage.getItem( ANSWER_MAP_TEMPLATE ) === null ) {
        do {
            questionText = prompt( "Form Question Text. Leave blank to save & quit.", "" ).toUpperCase();

            if ( questionText ) {
                let answer = prompt( `Default Answer to Question Text [${questionText}]. Case Sensitive.`, "" ),
                    index = answerMap.findIndex( ( element ) => element.QuestionText === questionText );

                index !== -1 ? answerMap[index].Answer = answer : alert( "Invalid Question Text, please try again." );
            }
        } while ( questionText );

        localStorage.setItem( ANSWER_MAP_TEMPLATE, JSON.stringify( answerMap ) );
    }

    if ( localStorage.getItem( ANSWER_MAP_TEMPLATE ) !== null ) {
        // Set new state
        localStorage.setItem( ACTIVE_SECTION, "rd39d9e0dc25548128a3b6604e4bd6207" ); // Main Form
        localStorage.setItem( ANSWER_MAP, localStorage.getItem( ANSWER_MAP_TEMPLATE ) );
    }
})();
