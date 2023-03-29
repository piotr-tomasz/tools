def header_fit(column_nb):
  '''Function intended to fit header to dataframe'''
  if column_nb == 14:
    head = ['Sample Nb.,[-]','Measurement time,[s]','Pressure-1,[kPa]','Temperature-1,[0C]','LPM-1,[l/min]','SLPM-1,[l/min]','Pressure-2,[kPa]','Temperature-2,[0C]',
    'LPM-2,[l/min]','SLPM-2,[l/min]','Gas type,[-]','Pt1,[kPa]','Pt2,[kPa]','Pt3,[kPa]']
  elif column_nb == 10:
    head = ['Sample Nb.,[-]','Measurement time,[s]','Pressure,[kPa]','Temperature,[0C]','LPM,[l/min]','SLPM,[l/min]','Gas type,[-]','Pt1,[kPa]','Pt2,[kPa]','Pt3,[kPa]']
  elif column_nb == 7:
    head = ['Sample Nb.,[-]','Measurement time,[s]','Pressure,[kPa]','Temperature,[0C]','LPM,[l/min]','SLPM,[l/min]','Gas type,[-]']
  else:
    print('Wrong number of columns')
    head = None
  return head
