import sys
import datetime
print(sys.executable)
print(sys.version)
print("helloworld")



print('Timestamp: {:%Y-%m-%d %H:%M:%S}'.format(datetime.datetime.now()))
print('Timestamp: {:%Y-%b-%d %H:%M:%S}'.format(datetime.datetime.now()))
print('Date now: %s' % datetime.datetime.now())
print('Date today: %s' % datetime.date.today())
