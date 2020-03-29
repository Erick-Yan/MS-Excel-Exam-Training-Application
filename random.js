$(document).ready(function(){

        var question;
        var unit;

        let all = new Array();
        all[0] = 'Open the HelpLess Airlines Passenger Numbers workbook. Share the workbook so that change history is saved for 90 days. Update the changes automatically every 10 minutes. Save the workbook when prompted to do so';
        all[1] = 'In the Numbers worksheet (HelpLess Airlines Passenger Numbers) modify the text in cell A2 to read Consolidated Passenger Data. Change the Font colour of the text in cell A2 to Red. Save the workbook. Now display all of the changes that have ever been made by anyone to this workbook. Don\’t highlight the changes on the screen and list the changes on a new worksheet';
        all[2] = '(HelpLess Airlines Passenger Numbers : Numbers) Configure Excel to enable background checking and display detected formula errors in Red. Leave the workbook open';
        all[3] = 'Open the Sales Information Data workbook. Show all formulas in the Sales Data worksheet. Now hide the formulas';
        all[4] = 'In the Sales Data worksheet (Sales Information Data) create a macro that formats the Row Height to be 30 points and changes the font size to 12. Name the macro Height and store it in this workbook. Run this macro on the cell range G2 to G10';
        all[5] = 'In the Sales Data (Sales Information Data) worksheet create a macro that applies a Currency number format and a Blue Gradient Data Bar rule. Name the macro Formatting and store it in this workbook. Apply the macro to the data in the cell range G12 to G20';
        all[6] = 'In the Sales Information Data workbook in the Buttons worksheet modify the Spin Button values to be between 1 and 50 and change in increments of 1';
        all[7] = 'In the Sales Information Data workbook protect the Sales Data worksheet with the password Hendrix';
        all[8] = 'In the Buttons worksheet (Sales Information Data) create a Button Form Control named Set Data Bars. Place it at cell A1. Assign the Formatting macro to the button and run it in cell A16';
        all[9] = 'Insert a Table in the Sales Data worksheet (Sales Information Data). The range to be selected is from A1 to G2009. The table should have headers. Name the Table Sales_Data. Include a Total Row. (This element is part of Core Excel and you should already know how to do this!) In cell I2010 create a formula which will total all of the data from G2 to G2009 using a fully qualified structured reference';
        all[10] = '(Sales Information Data : Sales Data) Find a list of all of the commands not currently displayed in the ribbon. Create a new Tab on the Ribbon and name it as MyNAmeTab replacing MyName with your first name i.e. DaveTab. Add two commands which are currently hidden into your new tab. Remove the Custom Tab you just created';
        all[11] = 'Change the Formula calculation options in the Sales Information Data workbook so that Iterative Calculations are enabled with a maximum iteration count of 50 and a maximum change of 0.002';
        all[12] = 'Change the Formula Calculation options so that Error checking does not apply to empty cells';
        all[13] = 'Remove the HelpLess Airlines workbook from Shared Use';
        all[14] = 'Create a macro in the Sales Information Data workbook which changes the colour of cell F1 to Green. Name the macro Green and have it apply to this workbook only. Copy the macro named Green to the HelpLess Airlines Passenger Numbers workbook';
        all[15] = 'Save the Sales Information Data workbook as a Template to any location of your choice';
        all[16] = 'In cell M1 of the Sales Information Data workbook create a formula which links to the Product Items worksheet in the HelpLess Airlines Passenger Numbers workbook. The formula should retrieve the total Price of the items as calculated in cell C10 of the Product Items worksheet but should not be a direct link to that cell';
        all[17] = 'Mark the HelpLess Airlines Passenger Numbers workbook as Final. Protect the structure of the workbook but don\’t provide a password to do so';

        class qu{
                constructor(quest, un){
                        this._quest = quest;
                        this._un = un;
                        this._array = 0;
                }

                quest1(){
                        this._array ++;
                }
        }

        function getNewQuote(){
                var n = all.length+1;
                var allNumber = Math.floor(n*Math.random());
                var q = all[allNumber];

                if(allNumber <= 17){
                        u = '(1) Manage Workbook (Review)';
                }
                const Question = new qu(q, u);
                question = Question._quest;
                unit = Question._un;
                Question.quest1();
                $('.Question').text(question);
                if(unit){
                        $('#Unit').text(unit);
                }else{
                        $('#Unit').text('Unknown');
                        }

                getNewQuote();
        }

        $('.button').on('click', function(event){
                event.preventDefault();
                if (all.length > 1){
                        all.pop();
                        getNewQuote();
                } else{
                        $('.Question').text('');
                        $('#Unit').text('Complete!');
                }

        });
});

/*let allConcepts = ['Macros','PivotTable','PivotChart'];
let tL  = 0;
let speed = 50;

function typeWriter(){
        for(let i = 0; i < 1; i--){
                let txt = allConcepts[Math.floor(Math.random*4)];
                if(tL < txt.length){
                        document.getElementById("Concepts") += txt.charAt(tL);
                        tL += 1;
                        setTimeout(typeWriter, speed);
                } else if(tL === txtLength){
                        document.getElementById("Concepts") -= txt.charAt(txt.Length);
                }
        }

}*/

const TypeWriter = function(txtElement, words, wait = 3000){
        this.txtElement = txtElement;
        this.words = words;
        this.txt = ' ';
        this.wordIndex = 0;
        this.wait = parseInt(wait, 10);
        this.type();
        this.isDeleting = false;
}

// Type Method
TypeWriter.prototype.type = function(){
        // Current index of words
        const current = this.wordIndex % this.words.length;
        // Get full text of current words
        const fullTxt = this.words[current];
        // Check if deleting
        if(this.isDeleting){
                // Remove char
                this.txt = fullTxt.substring(0, this.txt.length - 1);
        } else {
                //Add char
                this.txt = fullTxt.substring(0, this.txt.length + 1);
        }

        // Insert txt into element
        this.txtElement.innerHTML = `<span class='txt'> ${this.txt} <span>`;


        // Initial Type speed
        let typeSpeed = 300;

        if (this.isDeleting){
                typeSpeed /= 2;
        }

        // If word is complete
        if (!this.isDeleting && this.txt === fullTxt){
                // Pause at the end
                typeSpeed = this.wait;
                //Set delete to true
                this.isDeleting = true;
        } else if (this.isDeleting && this.txt === ''){
                this.isDeleting = false;
                // Move to next word
                this.wordIndex++;
                // Pause before start typing
                typeSpeed = 500;
        }

        setTimeout(()=> this.type(), typeSpeed);
}

// Init on DOM Load
document.addEventListener('DOMContentLoaded', init);
// Init Apply
function init(){
        const txtElement = document.querySelector('.txt-type');
        const words = JSON.parse(txtElement.getAttribute('data-words'));
        const wait = txtElement.getAttribute('data-wait');
        // Init TypeWriter
        new TypeWriter(txtElement,words, wait);
}

//ES6 class below
/*class TypeWriter{
        constructor(txtElement, words, wait = 6000){
                this.txtElement = txtElement;
                this.words = words;
                this.txt = ' ';
                this.wordIndex = 0;
                this.wait = parseInt(wait, 10);
                this.type();
                this.isDeleting = false;
        }
        type(){
                // Current index of words
                const current = this.wordIndex % this.words.length;
                // Get full text of current words
                const fullTxt = this.words[current];
                // Check if deleting
                if(this.isDeleting){
                        // Remove char
                        this.txt = fullTxt.substring(0, this.txt.length - 1);
                } else {
                        //Add char
                        this.txt = fullTxt.substring(0, this.txt.length + 1);
                }

                // Insert txt into element
                this.txtElement.innerHTML = `<span class='txt'> ${this.txt} <span>`;

                // Initial Type speed
                let typeSpeed = 300;

                if (this.isDeleting){
                        typeSpeed /= 2;
                }

                // If word is complete
                if (!this.isDeleting && this.txt === fullTxt){
                        // Pause at the end
                        typeSpeed = this.wait;
                        //Set delete to true
                        this.isDeleting = true;
                } else if (this.isDeleting && this.txt === ''){
                        this.isDeleting = false;
                        // Move to next word
                        this.wordIndex++;
                        // Pause before start typing
                        typeSpeed = 500;
                }

                setTimeout(()=> this.type(), typeSpeed);
        }
}

// Init on DOM Load
document.addEventListener('DOMContentLoaded', init);
// Init Apply
function init(){
        const txtElement = document.querySelector('.txt-type');
        const words = JSON.parse(txtElement.getAttribute('data-words'));
        const wait = txtElement.getAttribute('data-wait');
        // Init TypeWriter
        new TypeWriter(txtElement,words, wait);
}*/
