from flask import Flask, Response, render_template, request, redirect, url_for, send_file, session
from flask_wtf import FlaskForm
from wtforms import MultipleFileField, SubmitField
from werkzeug import secure_filename
import os
import pipe_design
import pipe_velocity
import gutter_spread
import wtforms.validators

app = Flask(__name__)

SECRET_KEY = os.urandom(32)
app.config['SECRET_KEY'] = SECRET_KEY

class network_upload(FlaskForm):
    design_files = MultipleFileField('Select Pipe Design Files')
    design_submit = SubmitField('Format Pipe Design Files')

    velocity_files = MultipleFileField('Select Pipe Velocity Files')
    velocity_submit = SubmitField('Format Pipe Velocity Files')

    spread_files = MultipleFileField('Select Gutter Spread Files') 
    spread_submit = SubmitField('Format Gutter Spread Files')

@app.route('/', methods=['GET', 'POST'])
def index():
    form = network_upload()
    if form.validate_on_submit(): 

        if form.design_submit.data:
            design_files = form.design_files.data
            design_names = [secure_filename(file.filename) for file in design_files]
            output_file = pipe_design.main(design_names, design_files)     
            return output_file

        if form.velocity_submit.data:
            velocity_files = form.velocity_files.data
            velocity_names = [secure_filename(file.filename) for file in velocity_files]
            output_file = pipe_velocity.main(velocity_names, velocity_files)     
            return output_file

        if form.spread_submit.data:
            spread_files = form.spread_files.data
            spread_names = [secure_filename(file.filename) for file in spread_files]
            output_file = gutter_spread.main(spread_names, spread_files)     
            return output_file

    return render_template('index.html', form=form)

@app.route('/design_report')
def design_report():
    return send_file('static/Design.rpt', mimetype='text/plain', 
    attachment_filename='Pipe Design.rpt', as_attachment=True)
    return redirect(url_for('index'))
	
@app.route('/velocity_report')
def velocity_report():
    return send_file('static/Velocity.rpt', mimetype='text/plain', 
    attachment_filename='Pipe Velocity.rpt', as_attachment=True)
    return redirect(url_for('index'))

@app.route('/spread_report')
def spread_report():
    return send_file('static/Spread.rpt', mimetype='text/plain', 
    attachment_filename='Gutter Spread.rpt', as_attachment=True)
    return redirect(url_for('index'))

@app.errorhandler(404)
def page_not_found(error):
    return render_template('404.html')

@app.errorhandler(500)
def internal_error(error):
    return render_template('500.html')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=1000, debug=True)
