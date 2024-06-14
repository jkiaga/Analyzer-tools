from flask_wtf import FlaskForm
from wtforms import SubmitField, DateField
from wtforms.validators import DataRequired

class DateForm(FlaskForm):
    start_date = DateField('Start Date', format='%Y-%m-%d', validators=[DataRequired()])
    end_date = DateField('End Date', format='%Y-%m-%d', validators=[DataRequired()])
    submit = SubmitField('Export Report')
