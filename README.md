# VBA of Wall Street

## Overview of Project
The purpose of this analysis was to provide Steve with a tool to help him advise his parents on which stocks they should invest in. The macro we built analyzes stock data for 12 companies for 2017 and 2018. We specifically analyzed volume and prices (starting/ending), to determine stock performance.

## Results
### Stock Performance
- Based on the 2017-2018 daily volume and returns, Steve should recommend his parents invest in ENPH and RUN stocks. Of the 12 stock analyzed, these two had positive returns both years and were the top two for daily volume in 2018. Since both of these stocks are peforming well, Steve should recommend investing in both stocks to diversify their portfolio. The market appears to be volatile given that of the 12 stocks, 9 or 75% went from postive returns to negative.

### Macro Performance
- Refactoring the code reduced the run time from .28 to .08 seconds which is around a 71% decrease. This perfomance boost came form utilizing arrays which allowed the macro to run through the 3000 rows one time and save all the information to the corresponding arrays -- whereas the original macro looped through the 3000 rows 12 times.

## Summary
- The advantages of refactoring is that it give you the opportunity to not only clean up your code but also make it more concise and efficient. Writing code is similar to drafting an essay. Although your first draft can get the job done, it most likely has errors and might not be concise. Through the editing process, you are able to correct your mistakes and reword paragraphs to create a more cohesive and stronger essay. Refactoring gives you the same opportunity for your code. Reviewing your code allows you to think about different ways to solve the problem and see if you can find more effiecient approaches. For example, refactoring our macro reduced the runtime by around 71%.

- A disadvantage of refactoring code is that it can be a lenghthy or time consuming process with diminishing returns. Although reducing our stock analysis run time by 70% is impressive, the original macro still only took less than a second to run. However, reducing a macro with an hour runtime by 70% would be more beneficial and much more noticeable. Going back to my essay comparison, spending hours editing a doctoral thesis makes sense whereas spending hours editing a single paragraph does not.
