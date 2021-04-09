import sys
import importlib
import traceback
import logging

test = importlib.import_module(sys.argv[1])
#print(("{}").format(test))
#logging.basicConfig(level=logging.INFO)
logging.basicConfig(level=logging.INFO, file='info.log')
if __name__ == '__main__':
    data = {}
    
    try:
        # Start the Main Sequence and wait until complete
        test.main(data)
    except Exception, e:
        traceback.print_exc()
    finally:
        traceback.print_exc()
        # Start the Cleanup Sequence and pass it the data object
        #test.cleanup(data)
