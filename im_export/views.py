import json
import logging
from django.http import HttpResponse, JsonResponse
from rest_framework.decorators import api_view, permission_classes
from rest_framework.permissions import IsAuthenticated
from rest_framework.response import Response
from insuree.apps import InsureeConfig
from .services import InsureeImportExportService

logger = logging.getLogger(__name__)


def check_user_rights(rights):
    class UserWithRights(IsAuthenticated):
        def has_permission(self, request, view):
            return super().has_permission(request, view) and request.user.has_perms(rights)

    return UserWithRights


@api_view(["POST"])
@permission_classes([check_user_rights(InsureeConfig.gql_mutation_create_insurees_perms, )])
def import_insurees(request):
    try:
        import_file = request.FILES.get('file', None)
        user = request.user
        dry_run = json.loads(request.POST.get('dry_run', 'false'))
        strategy = request.POST.get('strategy', InsureeImportExportService.Strategy.INSERT)

        success, totals, errors = InsureeImportExportService(user) \
            .import_insurees(import_file, dry_run=dry_run, strategy=strategy)
        return JsonResponse(data={'success': success, 'data': totals, 'errors': errors})
    except ValueError as e:
        logger.error("Error while importing insurees", exc_info=e)
        return Response({'success': False, 'error': str(e)}, status=400)
    except Exception as e:
        logger.error("Unexpected error while importing insurees", exc_info=e)
        return Response({'success': False, 'error': str(e)}, status=500)


@api_view(["GET"])
@permission_classes([check_user_rights(InsureeConfig.gql_query_insurees_perms, )])
def export_insurees(request):
    try:
        # the insurees the user may see only (InsureeResource.export: Insuree.get_queryset)
        export_format = request.GET.get("file_format", "csv")
        user = request.user

        content_type, export = InsureeImportExportService(user).export_insurees(export_format)
        response = HttpResponse(export, content_type=content_type)
        response['Content-Disposition'] = f'attachment; filename="insurees.{export_format}"'
        return response
    except ValueError as e:
        logger.error("Error while exporting insurees", exc_info=e)
        return Response({'success': False, 'error': str(e)}, status=400)
    except Exception as e:
        logger.error("Unexpected error while exporting insurees", exc_info=e)
        return Response({'success': False, 'error': str(e)}, status=500)
