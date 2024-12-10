from dotenv import load_dotenv
load_dotenv()

import looprunner
import traceback

if __name__ == '__main__':
    try:
        looprunner.run_loop()
    except Exception as err:
        print(traceback.format_exc())


# TODO This gives us the COMPLETE dials
# SELECT COUNT([HisCallNumber])
#   FROM [Voxco_Project_91722].[dbo].[Historic]
#   WHERE HisCallDate >= '2024-07-29 11:00:00.000' and HisCallDate <= '2024-07-30 11:00:00.000' and A4SCallNumber IS NULL and HisResult = 'CO'

# TODO Figure out why THIS matches bottom table
# SELECT COUNT([HisCallNumber])
#   FROM [Voxco_Project_91722].[dbo].[Historic]
#   WHERE HisCallDate >= '2024-07-29 11:00:00.000' and HisCallDate <= '2024-07-30 11:00:00.000' and A4SCallNumber IS NULL and HisResult NOT LIKE 'OK'