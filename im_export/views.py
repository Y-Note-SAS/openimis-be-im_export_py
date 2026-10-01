import json
import logging
from django.http import HttpResponse, JsonResponse
from rest_framework.decorators import api_view, permission_classes
from rest_framework.permissions import IsAuthenticated
from rest_framework.response import Response
from insuree.apps import InsureeConfig
from .services import InsureeImportExportService, BankImportService, FamilyImportExportService
from rest_framework import status
from .serializers import BankImportUploadSerializer
from .models import BankImport 
from core.models import InteractiveUser
from invoice.models import InvoiceEvent

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

        # success, totals, errors = InsureeImportExportService(user) \
        #     .import_insurees(import_file, dry_run=dry_run, strategy=strategy)
        # return JsonResponse(data={'success': success, 'data': totals, 'errors': errors})
        success, totals, errors = FamilyImportExportService(user) \
            .import_families(import_file, dry_run=dry_run, strategy=strategy)
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
        # TODO add location based filtering
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


@api_view(["POST"])
def import_exim_bank(request):
    serializer = BankImportUploadSerializer(data=request.data)
    serializer.is_valid(raise_exception=True)

    file = serializer.validated_data.get("file")
    user = InteractiveUser.objects.filter(login_name=request.user.username).first()
    service = BankImportService(user)
    errors = []
    successful_transactions = []
    transactions_result = {}

    try:
        logger.info(f"Upload fichier banque (user={user.id}, file={file})")

        bank_import = BankImport.objects.create(user=user, stored_file=file)

        transactions_result = service.parse_excel_exim(bank_import.stored_file)
            
        for idx, tx in enumerate(transactions_result["transactions"], start=1):
            try:
                results = service.reconcile_bank_transaction(tx)
                if not isinstance(results, list):
                    results = [results]
                for result in results:
                    successful_transactions.append({
                        "ligne": idx,
                        "insuree_chf_id": tx["insuree_chf_id"],
                        "amount": result["amount"],
                        "invoice_code": result["invoice_code"],
                        "payment_id": result["payment_id"],
                        "premium_uuid": result.get("premium_uuid"),
                        "status": result["status"],
                        "complete": result.get("detail_payment_id", None) is not None
                    })
            except Exception as exc:
                # Ajoute l'erreur avec l'index ou info utile pour retrouver la ligne
                errors.append(f"Ligne {idx}: {str(exc)}")
                logger.warning(f"Erreur sur la ligne {idx}: {exc}")
                # Si on trouve une facture liée, on logue l'erreur en event
                chfid = tx.get("insuree_chf_id", "").strip()
                msg = (
                    f"Erreur lors du traitement de la transaction ligne {idx} : {type(exc).__name__} - {str(exc)}. "
                    f"CHFID='{chfid}', Montant='{tx.get('amount_received', 'N/A')}', "
                    f"Date paiement='{tx.get('date_payment', 'N/A')}', Référence='{tx.get('code_ext', 'N/A')}'."
                )
                try:
                    invoice = service.find_invoice(chfid)
                    if invoice:
                        service.log_invoice_event(
                            user=user,
                            invoice=invoice,
                            event_type=InvoiceEvent.EventType.PAYMENT_ERROR,
                            message=msg
                        )
                except Exception as e:
                    logger.warning(f"Erreur lors du logging d'événement pour CHFID {chfid}: {e}")

        logger.info(f"{transactions_result['count']} transactions créditées extraites")

    except Exception as exc:
        logger.exception("Erreur durant l'import EXIM")
        errors.append(str(exc))

    return Response({
        "success": len(errors) == 0,
        "errors": errors,
        "processed": successful_transactions,
        **transactions_result
    }, status=status.HTTP_200_OK if not errors else status.HTTP_400_BAD_REQUEST)

@api_view(["POST"])
def import_bdc_bank(request):
    serializer = BankImportUploadSerializer(data=request.data)
    serializer.is_valid(raise_exception=True)

    file = serializer.validated_data.get("file")
    user = InteractiveUser.objects.filter(login_name=request.user.username).first()
    service = BankImportService(user)
    errors = []
    successful_transactions = []
    transactions_result = {}

    try:
        logger.info(f"Upload fichier banque BDC (user={user.id}, file={file})")

        bank_import = BankImport.objects.create(user=user, stored_file=file)

        transactions_result = service.parse_excel_bdc(bank_import.stored_file)
        
        for idx, tx in enumerate(transactions_result["transactions"], start=1):
            try:
                results = service.reconcile_bank_transaction(tx)
                if not isinstance(results, list):
                    results = [results]
                for result in results:
                    successful_transactions.append({
                        "ligne": idx,
                        "insuree_chf_id": tx["insuree_chf_id"],
                        "amount": result["amount"],
                        "invoice_code": result["invoice_code"],
                        "payment_id": result["payment_id"],
                        "premium_uuid": result.get("premium_uuid"),
                        "status": result["status"],
                        "complete": result.get("detail_payment_id", None) is not None
                    })
            except Exception as exc:
                chfid = tx.get("insuree_chf_id", "").strip()
                
                msg = (
                    f"Erreur lors du traitement de la transaction ligne {idx} : {type(exc).__name__} - {str(exc)}. "
                    f"CHFID='{chfid}', Montant='{tx.get('amount_received', 'N/A')}', "
                    f"Date paiement='{tx.get('date_payment', 'N/A')}', Référence='{tx.get('code_ext', 'N/A')}'."
                ) 
                errors.append(msg)
                logger.warning(msg)
                
                try:
                    invoice = service.find_invoice(chfid)
                    if invoice:
                        service.log_invoice_event(
                            user=user,
                            invoice=invoice,
                            event_type=InvoiceEvent.EventType.PAYMENT_ERROR,
                            message=msg
                        )
                except Exception as e:
                    logger.warning(f"Erreur lors du logging d'événement pour CHFID {chfid}: {e}")


        logger.info(f"{transactions_result['count']} transactions créditées extraites")

    except Exception as exc:
        logger.exception("Erreur durant l'import BDC")
        errors.append(str(exc))

    return Response({
        "success": len(errors) == 0,
        "errors": errors,
        "processed": successful_transactions,
        **transactions_result
    }, status=status.HTTP_200_OK if not errors else status.HTTP_400_BAD_REQUEST)


@api_view(["POST"])
def import_other_payment_method(request):
    serializer = BankImportUploadSerializer(data=request.data)
    serializer.is_valid(raise_exception=True)

    file = serializer.validated_data.get("file")
    user = InteractiveUser.objects.filter(login_name=request.user.username).first()
    service = BankImportService(user)
    errors = []
    successful_transactions = []
    transactions_result = {}

    try:
        logger.info(f"Upload fichier autre moyen de paiement (user={user.id}, file={file})")

        bank_import = BankImport.objects.create(user=user, stored_file=file)

        transactions_result = service.parse_excel_other_payment(bank_import.stored_file)
            
        for idx, tx in enumerate(transactions_result["transactions"], start=1):
            try:
                results = service.reconcile_bank_transaction(tx)
                if not isinstance(results, list):
                    results = [results]
                for result in results:
                    successful_transactions.append({
                        "ligne": idx,
                        "insuree_chf_id": tx["insuree_chf_id"],
                        "amount": result["amount"],
                        "invoice_code": result["invoice_code"],
                        "payment_id": result["payment_id"],
                        "premium_uuid": result.get("premium_uuid"),
                        "status": result["status"],
                        "complete": result.get("detail_payment_id", None) is not None
                    })
            except Exception as exc:
                # Ajoute l'erreur avec l'index ou info utile pour retrouver la ligne
                errors.append(f"Ligne {idx}: {str(exc)}")
                logger.warning(f"Erreur sur la ligne {idx}: {exc}")
                # Si on trouve une facture liée, on logue l'erreur en event
                chfid = tx.get("insuree_chf_id", "").strip()
                msg = (
                    f"Erreur lors du traitement de la transaction ligne {idx} : {type(exc).__name__} - {str(exc)}. "
                    f"CHFID='{chfid}', Montant='{tx.get('amount_received', 'N/A')}', "
                    f"Date paiement='{tx.get('date_payment', 'N/A')}', Référence='{tx.get('code_ext', 'N/A')}'."
                )
                try:
                    invoice = service.find_invoice(chfid)
                    if invoice:
                        service.log_invoice_event(
                            user=user,
                            invoice=invoice,
                            event_type=InvoiceEvent.EventType.PAYMENT_ERROR,
                            message=msg
                        )
                except Exception as e:
                    logger.warning(f"Erreur lors du logging d'événement pour CHFID {chfid}: {e}")

        logger.info(f"{transactions_result['count']} transactions créditées extraites")

    except Exception as exc:
        logger.exception("Erreur durant l'import EXIM")
        errors.append(str(exc))

    return Response({
        "success": len(errors) == 0,
        "errors": errors,
        "processed": successful_transactions,
        **transactions_result
    }, status=status.HTTP_200_OK if not errors else status.HTTP_400_BAD_REQUEST)