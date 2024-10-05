<h1>Fill in the Blank Quiz Program</h1>
    <p>A Python program designed for students to study using their given terms from an Excel file. The program presents questions by randomly selecting terms and asking the user to fill in the blank.</p>

  <h2>How it Works</h2>
    <ol>
        <li><strong>Prepare the Excel file (<code>terms.xlsx</code>)</strong>
            <ul>
                <li>Ensure the <code>terms.xlsx</code> file contains terms in <strong>Column A</strong>.</li>
                <li>You can have any number of terms (one per row).</li>
                <li><strong>Important</strong>: Close the Excel file before starting the quiz.</li>
            </ul>
        </li>
        <li><strong>Running the Quiz</strong>
            <ul>
                <li>The program will ask how many questions you’d like to answer.</li>
                <li>The number of questions you enter must be <strong>less than or equal to</strong> the total number of available terms.</li>
                <li>The quiz will then start, displaying one fill-in-the-blank question at a time.</li>
                <li>After the last question is answered, your score will be displayed along with the total number of questions.</li>
            </ul>
        </li>
    </ol>

  <h2>Logic Behind the Program</h2>
    <ul>
        <li><strong>Term Selection</strong>: Only words with more than <strong>3 characters</strong> are selected to be blanked out, ensuring that important words are chosen for the quiz.</li>
        <li><strong>Answer Handling</strong>: 
            <ul>
                <li>The answer comparison is <strong>not case-sensitive</strong>.</li>
                <li>Hyphens or other punctuation in the user's answer are ignored, allowing for more flexible input.</li>
            </ul>
        </li>
        <li><strong>Error Handling</strong>: 
            <ul>
                <li>The program ensures that any errors related to missing dependencies (such as missing modules) are handled with clear instructions.</li>
                <li>If the Excel file is open, the program will prompt you to close it before proceeding.</li>
            </ul>
        </li>
    </ul>

  <h2>Features</h2>
    <ul>
        <li>Randomly selects terms from the Excel file and generates fill-in-the-blank questions.</li>
        <li>Ensures blanked-out words are meaningful (more than 3 characters).</li>
        <li>Case-insensitive and punctuation-tolerant answers.</li>
        <li>Displays the user's score at the end of the quiz.</li>
    </ul>

  <h2>Getting Started</h2>
    <ol>
        <li><strong>Install the dependencies</strong>:
            <pre><code>pip install openpyxl</code></pre>
        </li>
        <li><strong>Place your terms in <code>terms.xlsx</code></strong>:
            <ul>
                <li>Terms must be placed in Column A of the Excel sheet.</li>
            </ul>
        </li>
        <li><strong>Run the program</strong>:
            <p>Execute the Python script, follow the prompts, and enjoy the quiz!</p>
        </li>
    </ol>

  <h2>Example</h2>
    <p>Here’s what a typical quiz question looks like:</p>
    <pre><code>Question 1: Fill in the blank: The _____ is blue.
Your answer: sky
Correct!</code></pre>
