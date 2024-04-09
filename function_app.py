"""Entrypoint for the Azure Functions app."""

import http

from azure import functions

from ctk_functions.functions.file_conversion import (
    controller as file_conversion_controller,
)
from ctk_functions.functions.intake import controller as intake_controller

app = functions.FunctionApp()


@app.function_name(name="get-intake-report")
@app.route(route="intake-report/{survey_id}", auth_level=functions.AuthLevel.FUNCTION)
async def get_intake_report(req: functions.HttpRequest) -> functions.HttpResponse:
    """Generates an intake report for a survey.

    Args:
        req: The HTTP request object.

    Returns:
        The HTTP response containing the .docx file.
    """
    survey_id = req.route_params.get("survey_id")
    if not survey_id:
        return functions.HttpResponse(
            "Please provide a survey ID.", status_code=http.HTTPStatus.BAD_REQUEST
        )

    docx_bytes = intake_controller.get_intake_report(survey_id)
    return functions.HttpResponse(
        body=docx_bytes,
        status_code=http.HTTPStatus.OK,
    )


@app.function_name(name="markdown2docx")
@app.route(route="markdown2docx", auth_level=functions.AuthLevel.FUNCTION)
async def markdown2docx(req: functions.HttpRequest) -> functions.HttpResponse:
    """Converts a Markdown document to a .docx file.

    Args:
        req: The HTTP request object.

    Returns:
        The HTTP response containing the .docx file.
    """
    markdown = req.get_body().decode("utf-8")
    if not markdown:
        return functions.HttpResponse(
            "Please provide a Markdown document.",
            status_code=http.HTTPStatus.BAD_REQUEST,
        )
    docx_bytes = file_conversion_controller.markdown2docx(markdown)
    return functions.HttpResponse(
        body=docx_bytes,
        status_code=http.HTTPStatus.OK,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
